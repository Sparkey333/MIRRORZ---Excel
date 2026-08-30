import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  XmlReader,
  XmlToken,
  XmlWriter,
  decodeEntities,
  escapeSharedString,
  escapeXml,
  unescapeSharedString,
} from '../src/xml.js';
import { readZip } from '../src/zip.js';

const FIXTURES = new URL('../../../fixtures/generated/', import.meta.url);
const part = (file: string, name: string) =>
  new TextDecoder().decode(
    readZip(new Uint8Array(readFileSync(new URL(file, FIXTURES)))).get(name)!.data(),
  );

/** Collect a flat token trace, for compact structural assertions. */
function trace(xml: string): string[] {
  const r = new XmlReader(xml);
  const out: string[] = [];
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t === XmlToken.Open) out.push(`<${r.name}>`);
    else if (t === XmlToken.Close) out.push(`</${r.name}>`);
    else if (t === XmlToken.Text && r.text.trim()) out.push(`"${r.text}"`);
  }
  return out;
}

describe('structure', () => {
  it('walks nested elements', () => {
    expect(trace('<a><b>hi</b><c/></a>')).toEqual([
      '<a>',
      '<b>',
      '"hi"',
      '</b>',
      '<c>',
      '</c>',
      '</a>',
    ]);
  });

  it('reports a self-closing tag as an open immediately followed by a close', () => {
    const r = new XmlReader('<a/>');
    expect(r.next()).toBe(XmlToken.Open);
    expect(r.isSelfClosing).toBe(true);
    expect(r.next()).toBe(XmlToken.Close);
    expect(r.next()).toBe(XmlToken.EOF);
  });

  it('tracks depth', () => {
    const r = new XmlReader('<a><b><c/></b></a>');
    const depths: number[] = [];
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Open) depths.push(r.depth);
    }
    expect(depths).toEqual([1, 2, 3]);
  });

  it('strips namespace prefixes into localName', () => {
    const r = new XmlReader('<x:worksheet xmlns:x="urn:x"><x:sheetData/></x:worksheet>');
    r.next();
    expect(r.name).toBe('x:worksheet');
    expect(r.localName).toBe('worksheet');
  });

  it('skips declarations, comments and CDATA appropriately', () => {
    const r = new XmlReader('<?xml version="1.0"?><!-- note --><a><![CDATA[<raw> & stuff]]></a>');
    expect(r.next()).toBe(XmlToken.ProcessingInstruction);
    expect(r.next()).toBe(XmlToken.Comment);
    expect(r.text).toBe(' note ');
    expect(r.next()).toBe(XmlToken.Open);
    expect(r.next()).toBe(XmlToken.CData);
    expect(r.text).toBe('<raw> & stuff');
  });

  it('does not expand entity definitions in a DOCTYPE', () => {
    // The billion-laughs shape: we must treat the DOCTYPE as opaque text.
    const bomb = '<!DOCTYPE lolz [<!ENTITY lol "lol">]><a>&lol;</a>';
    const r = new XmlReader(bomb);
    expect(r.next()).toBe(XmlToken.Doctype);
    r.next();
    // The undefined entity stays literal rather than expanding.
    expect(r.readText()).toBe('&lol;');
  });
});

describe('attributes', () => {
  it('reads by qualified and local name', () => {
    const r = new XmlReader('<c r="A1" s="3" t="s" xml:space="preserve"/>');
    r.next();
    expect(r.attr('r')).toBe('A1');
    expect(r.attr('s')).toBe('3');
    expect(r.attr('xml:space')).toBe('preserve');
    expect(r.attr('space')).toBe('preserve');
    expect(r.attr('missing')).toBeUndefined();
  });

  it('handles single quotes, spacing and empty values', () => {
    const r = new XmlReader("<a x='1'   y = \"2\" z=''/>");
    r.next();
    expect(r.attrs()).toEqual({ x: '1', y: '2', z: '' });
  });

  it('does not stop at a > inside an attribute value', () => {
    // Real OOXML: comparison operators appear in formula attributes.
    const r = new XmlReader('<f t="shared" ref="A1:A9" si="0" ca="1" x="a &gt; b"/>');
    r.next();
    expect(r.attr('x')).toBe('a > b');
    expect(r.attr('ref')).toBe('A1:A9');
  });

  it('decodes entities in attribute values', () => {
    const r = new XmlReader('<a v="&lt;tag&gt; &amp; &quot;q&quot; &apos;a&apos; &#65;&#x42;"/>');
    r.next();
    expect(r.attr('v')).toBe('<tag> & "q" \'a\' AB');
  });

  it('caches the attribute map', () => {
    const r = new XmlReader('<a x="1"/>');
    r.next();
    expect(r.attrs()).toBe(r.attrs());
  });
});

describe('text extraction', () => {
  it('concatenates across nested runs, as sharedStrings needs', () => {
    const r = new XmlReader('<si><r><t>Hello </t></r><r><t>world</t></r></si>');
    r.next();
    expect(r.readText()).toBe('Hello world');
  });

  it('returns empty for a self-closing element and lands on its close', () => {
    const r = new XmlReader('<a><b/><c>x</c></a>');
    r.next();
    r.next();
    expect(r.readText()).toBe('');
    expect(r.next()).toBe(XmlToken.Open);
    expect(r.name).toBe('c');
  });

  it('skipElement jumps the whole subtree', () => {
    const r = new XmlReader('<a><skip><deep><deeper/></deep></skip><after>x</after></a>');
    r.next();
    r.next();
    expect(r.name).toBe('skip');
    r.skipElement();
    expect(r.next()).toBe(XmlToken.Open);
    expect(r.name).toBe('after');
  });

  it('readRaw returns the original markup verbatim, for preserved parts', () => {
    const src = '<root><keep a="1"><child/>text</keep><next/></root>';
    const r = new XmlReader(src);
    r.next();
    r.next();
    expect(r.readRaw()).toBe('<keep a="1"><child/>text</keep>');
    expect(r.next()).toBe(XmlToken.Open);
    expect(r.name).toBe('next');
  });
});

describe('escaping', () => {
  it('escapes the five predefined entities', () => {
    expect(escapeXml('a<b>c&d"e\'f')).toBe('a&lt;b&gt;c&amp;d&quot;e&apos;f');
  });

  it('drops control characters XML 1.0 cannot represent', () => {
    expect(escapeXml('abc')).toBe('abc');
    expect(escapeXml('keep\ttab\nnewline\rcr')).toBe('keep\ttab\nnewline\rcr');
  });

  it('leaves clean text untouched without allocating', () => {
    const clean = 'nothing to escape here';
    expect(escapeXml(clean)).toBe(clean);
  });

  it.each([
    ['line1\nline2', '_x000A_'],
    ['tab\there', '_x0009_'],
    ['cr\rhere', '_x000D_'],
  ])('escapes control characters in shared strings: %s', (input, expected) => {
    expect(escapeSharedString(input)).toContain(expected);
    expect(unescapeSharedString(escapeSharedString(input))).toBe(input);
  });

  it('escapes an underscore that would read back as an escape', () => {
    expect(escapeSharedString('_x0041_')).toBe('_x005F_x0041_');
    expect(unescapeSharedString('_x005F_x0041_')).toBe('_x0041_');
  });

  it('leaves a harmless underscore alone', () => {
    expect(escapeSharedString('snake_case')).toBe('snake_case');
  });

  it('decodeEntities passes through text with no ampersand', () => {
    const s = 'plain';
    expect(decodeEntities(s)).toBe(s);
  });

  it('leaves an unknown entity literal rather than dropping it', () => {
    expect(decodeEntities('&nbsp;')).toBe('&nbsp;');
  });
});

describe('writer', () => {
  it('builds a document', () => {
    const w = new XmlWriter();
    w.open('worksheet', { xmlns: 'http://x' })
      .open('sheetData')
      .open('row', { r: 1 })
      .leaf('c', '5', { r: 'A1', t: 'n' })
      .close()
      .close()
      .empty('pageSetup', { orientation: 'portrait' })
      .close();
    expect(w.toString()).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<worksheet xmlns="http://x"><sheetData><row r="1"><c r="A1" t="n">5</c></row></sheetData>' +
        '<pageSetup orientation="portrait"/></worksheet>',
    );
  });

  it('omits undefined and false attributes but keeps zero', () => {
    const w = new XmlWriter(false);
    w.empty('a', { keep: 0, drop: undefined, off: false, on: true });
    expect(w.toString()).toBe('<a keep="0" on="1"/>');
  });

  it('refuses to serialise with elements left open', () => {
    const w = new XmlWriter(false);
    w.open('a');
    expect(() => w.toString()).toThrow(/left open/);
  });

  it('round-trips through the reader', () => {
    const w = new XmlWriter(false);
    w.open('a').leaf('b', 'x < y & z').close();
    const r = new XmlReader(w.toString());
    r.next();
    r.next();
    expect(r.readText()).toBe('x < y & z');
  });
});

describe('against real OOXML parts', () => {
  it('parses workbook.xml and finds every sheet', () => {
    const r = new XmlReader(part('features.xlsx', 'xl/workbook.xml'));
    const sheets: string[] = [];
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Open && r.localName === 'sheet') sheets.push(r.attr('name')!);
    }
    expect(sheets).toEqual(['Features', 'Charts', 'Comments', 'HiddenSheet']);
  });

  it('reads inline strings, the encoding openpyxl emits', () => {
    const r = new XmlReader(part('basic-types.xlsx', 'xl/worksheets/sheet1.xml'));
    const strings: string[] = [];
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Open && r.localName === 'is') strings.push(r.readText());
    }
    expect(strings).toContain('éàü 你好 \u{1f600}');
  });

  it('parses a worksheet and counts cells', () => {
    const r = new XmlReader(part('formulas.xlsx', 'xl/worksheets/sheet2.xml'));
    let cells = 0;
    let formulas = 0;
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Open) {
        if (r.localName === 'c') cells++;
        else if (r.localName === 'f') formulas++;
      }
    }
    expect(cells).toBeGreaterThan(300);
    expect(formulas).toBeGreaterThan(130);
  });

  it('reads sharedStrings including unicode and emoji', () => {
    // openpyxl writes inline strings; LibreOffice rewrites them into a shared
    // string table. The .calc.xlsx pass gives us a fixture for that encoding,
    // and our reader has to handle both.
    const r = new XmlReader(part('basic-types.calc.xlsx', 'xl/sharedStrings.xml'));
    const strings: string[] = [];
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Open && r.localName === 'si') strings.push(r.readText());
    }
    expect(strings).toContain('éàü 你好 \u{1f600}');
    expect(strings).toContain('he said "hi"');
  });

  it('streams a megabyte-scale sheet without a DOM', () => {
    const xml = part('large.xlsx', 'xl/worksheets/sheet1.xml');
    expect(xml.length).toBeGreaterThan(1_000_000);
    const started = performance.now();
    const r = new XmlReader(xml);
    let rows = 0;
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t === XmlToken.Open && r.localName === 'row') rows++;
    }
    const elapsed = performance.now() - started;
    expect(rows).toBe(50_001);
    // Generous bound: this is a smoke test against accidental quadratic work,
    // not a benchmark, so it should not go red on a loaded CI machine.
    expect(elapsed).toBeLessThan(5000);
  });
});
