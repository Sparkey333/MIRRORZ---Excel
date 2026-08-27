/**
 * Streaming XML pull-parser and writer, tuned for OOXML.
 *
 * A DOM is not an option here. A worksheet part in a real workbook routinely
 * runs to hundreds of megabytes, and materialising a node object per cell would
 * cost more memory than the sheet itself. This parser walks the document with a
 * cursor, allocating only for the pieces a caller actually asks for - attribute
 * objects are built lazily, so skipping an element costs nothing.
 *
 * It is deliberately not a general-purpose XML processor. It handles what OOXML
 * and ODF actually contain, and rejects the rest loudly rather than guessing.
 * In particular there is no DTD/entity-definition support, which also means the
 * classic "billion laughs" entity-expansion attack has no surface here - a real
 * consideration for an app that opens files from strangers.
 */

export enum XmlToken {
  /** `<tag ...>` */
  Open,
  /** `</tag>` */
  Close,
  /** `<tag ... />` - reported once, as an Open immediately followed by a Close. */
  SelfClose,
  Text,
  CData,
  Comment,
  ProcessingInstruction,
  Doctype,
  EOF,
}

export class XmlError extends Error {
  constructor(
    message: string,
    readonly offset: number,
  ) {
    super(`${message} (at offset ${offset})`);
    this.name = 'XmlError';
  }
}

/**
 * Cursor over an XML document.
 *
 * Usage is a loop over `next()`, testing `token` and `name`, then pulling
 * attributes or text only where needed.
 */
export class XmlReader {
  /** Current token kind. */
  token: XmlToken = XmlToken.EOF;
  /** Qualified element name for Open/Close/SelfClose, e.g. `a:t`. */
  name = '';
  /** Local name with any namespace prefix stripped, e.g. `t`. */
  localName = '';
  /** Text content for Text/CData tokens. Entity-decoded. */
  text = '';
  /** Nesting depth of the current element, root being 1. */
  depth = 0;

  private pos = 0;
  private attrStart = -1;
  private attrEnd = -1;
  private attrsCache: Record<string, string> | null = null;
  /** Set when a self-closing tag has emitted Open and still owes a Close. */
  private pendingClose = false;

  constructor(private readonly src: string) {}

  get offset(): number {
    return this.pos;
  }

  /**
   * Advance to the next token and return it. `XmlToken.EOF` ends the document.
   *
   * Returning the token rather than a boolean is deliberate: callers compare the
   * returned value, which keeps the type checker honest about the fact that this
   * call mutates `token`.
   */
  next(): XmlToken {
    if (this.pendingClose) {
      this.pendingClose = false;
      this.token = XmlToken.Close;
      this.depth--;
      this.attrsCache = null;
      this.attrStart = this.attrEnd = -1;
      return this.token;
    }
    this.attrsCache = null;

    const src = this.src;
    if (this.pos >= src.length) {
      this.token = XmlToken.EOF;
      return this.token;
    }

    if (src.charCodeAt(this.pos) !== 60 /* < */) {
      // Character data up to the next tag.
      const lt = src.indexOf('<', this.pos);
      const end = lt < 0 ? src.length : lt;
      this.text = decodeEntities(src.slice(this.pos, end));
      this.pos = end;
      this.token = XmlToken.Text;
      return this.token;
    }

    // Markup.
    const c1 = src.charCodeAt(this.pos + 1);
    if (c1 === 33 /* ! */) {
      if (src.startsWith('<!--', this.pos)) {
        const end = src.indexOf('-->', this.pos + 4);
        if (end < 0) throw new XmlError('unterminated comment', this.pos);
        this.text = src.slice(this.pos + 4, end);
        this.pos = end + 3;
        this.token = XmlToken.Comment;
        return this.token;
      }
      if (src.startsWith('<![CDATA[', this.pos)) {
        const end = src.indexOf(']]>', this.pos + 9);
        if (end < 0) throw new XmlError('unterminated CDATA section', this.pos);
        // CDATA is literal: no entity decoding.
        this.text = src.slice(this.pos + 9, end);
        this.pos = end + 3;
        this.token = XmlToken.CData;
        return this.token;
      }
      // DOCTYPE or another declaration. We skip it rather than interpret it;
      // in particular we never expand entity definitions, so a document that
      // declares them cannot use us to amplify itself.
      //
      // An internal subset - `<!DOCTYPE x [ ... ]>` - contains its own '>'
      // characters, so we must skip to the matching ']' before looking for the
      // declaration's real end.
      let end: number;
      const bracket = src.indexOf('[', this.pos);
      const firstGt = src.indexOf('>', this.pos);
      if (bracket >= 0 && (firstGt < 0 || bracket < firstGt)) {
        const closeBracket = src.indexOf(']', bracket);
        if (closeBracket < 0) throw new XmlError('unterminated internal DTD subset', this.pos);
        end = src.indexOf('>', closeBracket);
      } else {
        end = firstGt;
      }
      if (end < 0) throw new XmlError('unterminated declaration', this.pos);
      this.text = src.slice(this.pos + 2, end);
      this.pos = end + 1;
      this.token = XmlToken.Doctype;
      return this.token;
    }

    if (c1 === 63 /* ? */) {
      const end = src.indexOf('?>', this.pos + 2);
      if (end < 0) throw new XmlError('unterminated processing instruction', this.pos);
      this.text = src.slice(this.pos + 2, end);
      this.pos = end + 2;
      this.token = XmlToken.ProcessingInstruction;
      return this.token;
    }

    if (c1 === 47 /* / */) {
      const end = src.indexOf('>', this.pos + 2);
      if (end < 0) throw new XmlError('unterminated closing tag', this.pos);
      this.setName(src.slice(this.pos + 2, end).trim());
      this.pos = end + 1;
      this.depth--;
      this.token = XmlToken.Close;
      return this.token;
    }

    // Opening tag. Find its end, taking care not to stop at a '>' inside an
    // attribute value - legal XML, and OOXML formula attributes do contain them.
    let p = this.pos + 1;
    const nameStart = p;
    while (p < src.length) {
      const ch = src.charCodeAt(p);
      if (ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 47 || ch === 62) break;
      p++;
    }
    this.setName(src.slice(nameStart, p));

    this.attrStart = p;
    let quote = 0;
    while (p < src.length) {
      const ch = src.charCodeAt(p);
      if (quote) {
        if (ch === quote) quote = 0;
      } else if (ch === 34 || ch === 39) {
        quote = ch;
      } else if (ch === 62 /* > */) {
        break;
      }
      p++;
    }
    if (p >= src.length) throw new XmlError(`unterminated tag <${this.name}>`, this.pos);

    const selfClosing = src.charCodeAt(p - 1) === 47; /* / */
    this.attrEnd = selfClosing ? p - 1 : p;
    this.pos = p + 1;
    this.depth++;
    this.token = XmlToken.Open;
    if (selfClosing) this.pendingClose = true;
    return this.token;
  }

  private setName(qname: string): void {
    this.name = qname;
    const colon = qname.indexOf(':');
    this.localName = colon < 0 ? qname : qname.slice(colon + 1);
  }

  /** True when the current token opened an element that closes immediately. */
  get isSelfClosing(): boolean {
    return this.pendingClose;
  }

  /**
   * One attribute by qualified or local name, without building the whole map.
   * This is the hot path: a cell element is `<c r="A1" s="3" t="s">`, and
   * scanning for three known names beats allocating an object per cell.
   */
  attr(name: string): string | undefined {
    if (this.attrStart < 0) return undefined;
    const src = this.src;
    let p = this.attrStart;
    const end = this.attrEnd;
    while (p < end) {
      while (p < end && isSpace(src.charCodeAt(p))) p++;
      if (p >= end) break;
      const nameStart = p;
      while (p < end && src.charCodeAt(p) !== 61 /* = */ && !isSpace(src.charCodeAt(p))) p++;
      const rawName = src.slice(nameStart, p);
      while (p < end && isSpace(src.charCodeAt(p))) p++;
      if (src.charCodeAt(p) !== 61) {
        // A valueless attribute is not legal XML; skip it rather than derail.
        continue;
      }
      p++;
      while (p < end && isSpace(src.charCodeAt(p))) p++;
      const q = src.charCodeAt(p);
      if (q !== 34 && q !== 39) throw new XmlError('unquoted attribute value', p);
      p++;
      const valueStart = p;
      while (p < end && src.charCodeAt(p) !== q) p++;
      const value = src.slice(valueStart, p);
      p++;
      if (rawName === name || localOf(rawName) === name) return decodeEntities(value);
    }
    return undefined;
  }

  /** All attributes as a plain object, keyed by qualified name. */
  attrs(): Record<string, string> {
    if (this.attrsCache) return this.attrsCache;
    const out: Record<string, string> = {};
    if (this.attrStart >= 0) {
      const src = this.src;
      let p = this.attrStart;
      const end = this.attrEnd;
      while (p < end) {
        while (p < end && isSpace(src.charCodeAt(p))) p++;
        if (p >= end) break;
        const nameStart = p;
        while (p < end && src.charCodeAt(p) !== 61 && !isSpace(src.charCodeAt(p))) p++;
        const rawName = src.slice(nameStart, p);
        while (p < end && isSpace(src.charCodeAt(p))) p++;
        if (src.charCodeAt(p) !== 61) continue;
        p++;
        while (p < end && isSpace(src.charCodeAt(p))) p++;
        const q = src.charCodeAt(p);
        if (q !== 34 && q !== 39) throw new XmlError('unquoted attribute value', p);
        p++;
        const valueStart = p;
        while (p < end && src.charCodeAt(p) !== q) p++;
        out[rawName] = decodeEntities(src.slice(valueStart, p));
        p++;
      }
    }
    this.attrsCache = out;
    return out;
  }

  /**
   * Concatenated text of the element that is currently open, consuming through
   * its closing tag. Nested markup contributes its text but not its tags, which
   * is what `<is><t>` runs in sharedStrings need.
   */
  readText(): string {
    if (this.token !== XmlToken.Open) return '';
    if (this.pendingClose) {
      this.next();
      return '';
    }
    const target = this.depth - 1;
    let out = '';
    for (let t = this.next(); t !== XmlToken.EOF; t = this.next()) {
      if (t === XmlToken.Text || t === XmlToken.CData) {
        out += this.text;
      } else if (t === XmlToken.Close && this.depth === target) {
        break;
      }
    }
    return out;
  }

  /** Skip past the currently open element and everything inside it. */
  skipElement(): void {
    if (this.token !== XmlToken.Open) return;
    if (this.pendingClose) {
      this.next();
      return;
    }
    const target = this.depth - 1;
    for (let t = this.next(); t !== XmlToken.EOF; t = this.next()) {
      if (t === XmlToken.Close && this.depth === target) return;
    }
  }

  /**
   * Raw source of the currently open element, including its own tags.
   * Used to preserve parts we do not model yet, byte-for-byte.
   */
  readRaw(): string {
    if (this.token !== XmlToken.Open) return '';
    // Walk back to this tag's '<'. The reader has already consumed it, and
    // rescanning a short distance is cheaper than tracking a start offset for
    // every tag when almost none of them are preserved.
    const start = this.src.lastIndexOf('<', this.attrStart);
    if (this.pendingClose) {
      this.next();
      return this.src.slice(start, this.pos);
    }
    const target = this.depth - 1;
    for (let t = this.next(); t !== XmlToken.EOF; t = this.next()) {
      if (t === XmlToken.Close && this.depth === target) break;
    }
    return this.src.slice(start, this.pos);
  }
}

function isSpace(ch: number): boolean {
  return ch === 32 || ch === 9 || ch === 10 || ch === 13;
}

function localOf(qname: string): string {
  const i = qname.indexOf(':');
  return i < 0 ? qname : qname.slice(i + 1);
}

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/**
 * Decode the five predefined entities plus numeric character references.
 *
 * The fast path matters: most attribute values and text nodes contain no
 * ampersand at all, so we check for one before doing any work.
 */
export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match, body: string) => {
    if (body.charCodeAt(0) === 35 /* # */) {
      const code =
        body.charCodeAt(1) === 120 || body.charCodeAt(1) === 88
          ? Number.parseInt(body.slice(2), 16)
          : Number.parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10_ffff) return match;
      try {
        return String.fromCodePoint(code);
      } catch {
        return match;
      }
    }
    return NAMED_ENTITIES[body] ?? match;
  });
}

/**
 * Escape text for XML content or an attribute value.
 *
 * Beyond the standard five, this strips control characters that are simply not
 * representable in XML 1.0. Excel encodes those in strings as `_x000D_`-style
 * escapes rather than emitting them raw, and so do we (see `escapeSharedString`).
 */
export function escapeXml(s: string): string {
  let out = '';
  let last = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    let rep: string | undefined;
    if (c === 38) rep = '&amp;';
    else if (c === 60) rep = '&lt;';
    else if (c === 62) rep = '&gt;';
    else if (c === 34) rep = '&quot;';
    else if (c === 39) rep = '&apos;';
    else if (c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) rep = '';
    if (rep !== undefined) {
      out += s.slice(last, i) + rep;
      last = i + 1;
    }
  }
  return last === 0 ? s : out + s.slice(last);
}

/**
 * Excel's `_xHHHH_` escape for characters it will not write literally into a
 * shared string, and the literal underscore it must therefore also escape.
 */
export function escapeSharedString(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x20 || c === 0x7f) {
      out += `_x${c.toString(16).toUpperCase().padStart(4, '0')}_`;
    } else if (c === 0x5f /* _ */ && /^_x[0-9a-fA-F]{4}_/.test(s.slice(i))) {
      // Only escape an underscore that would otherwise be read back as an escape.
      out += '_x005F_';
    } else {
      out += s[i];
    }
  }
  return out;
}

export function unescapeSharedString(s: string): string {
  if (s.indexOf('_x') < 0) return s;
  return s.replace(/_x([0-9a-fA-F]{4})_/g, (match, hex: string) => {
    const code = Number.parseInt(hex, 16);
    // _x005F_ is the escaped underscore itself.
    return code === 0x5f ? '_' : String.fromCharCode(code);
  });
}

/** Minimal, allocation-conscious XML writer. */
export class XmlWriter {
  private readonly parts: string[] = [];
  private readonly stack: string[] = [];

  constructor(declaration = true, standalone = true) {
    if (declaration) {
      this.parts.push(
        `<?xml version="1.0" encoding="UTF-8" standalone="${standalone ? 'yes' : 'no'}"?>\r\n`,
      );
    }
  }

  open(name: string, attrs?: Record<string, string | number | boolean | undefined>): this {
    this.parts.push('<', name);
    this.writeAttrs(attrs);
    this.parts.push('>');
    this.stack.push(name);
    return this;
  }

  /** A `<tag .../>` with no children. */
  empty(name: string, attrs?: Record<string, string | number | boolean | undefined>): this {
    this.parts.push('<', name);
    this.writeAttrs(attrs);
    this.parts.push('/>');
    return this;
  }

  close(): this {
    const name = this.stack.pop();
    if (name === undefined) throw new Error('XmlWriter: close() with no open element');
    this.parts.push('</', name, '>');
    return this;
  }

  /** `<tag>text</tag>` in one call. */
  leaf(
    name: string,
    text: string | number | boolean,
    attrs?: Record<string, string | number | boolean | undefined>,
  ): this {
    this.parts.push('<', name);
    this.writeAttrs(attrs);
    this.parts.push('>', escapeXml(String(text)), '</', name, '>');
    return this;
  }

  text(value: string): this {
    this.parts.push(escapeXml(value));
    return this;
  }

  /** Insert pre-serialised XML verbatim, for preserved parts. */
  raw(xml: string): this {
    this.parts.push(xml);
    return this;
  }

  private writeAttrs(attrs?: Record<string, string | number | boolean | undefined>): void {
    if (!attrs) return;
    for (const key in attrs) {
      const v = attrs[key];
      if (v === undefined || v === false) continue;
      this.parts.push(' ', key, '="', escapeXml(v === true ? '1' : String(v)), '"');
    }
  }

  toString(): string {
    if (this.stack.length > 0) {
      throw new Error(`XmlWriter: ${this.stack.length} element(s) left open: ${this.stack.join(', ')}`);
    }
    return this.parts.join('');
  }

  toBytes(): Uint8Array {
    return new TextEncoder().encode(this.toString());
  }
}
