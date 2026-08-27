/**
 * Open Packaging Conventions: the container layer under every Office format.
 *
 * The design principle here is that the document model is a *projection* over
 * the package, never a replacement for it. Every zip entry is retained. Parts we
 * understand get parsed into typed models; parts we do not are kept as raw bytes
 * and written back untouched. That is the difference between a tool people trust
 * with their files and one whose own documentation has to warn that "images and
 * charts will be lost from existing files if they are opened and saved with the
 * same name".
 *
 * The same rule applies one level down, inside parts we *do* model: unrecognised
 * child elements are captured as opaque XML fragments and re-emitted in the
 * right schema position. `extLst` in particular is where Excel hides data bars,
 * several icon sets, sparklines, and slicer references, and dropping it quietly
 * degrades a workbook.
 */

import { XmlReader, XmlToken, XmlWriter } from './xml.js';
import { type ZipEntry, type ZipWriteEntry, readZip, writeZip } from './zip.js';

/** Relationship type URIs we care about, minus their long common prefix. */
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const REL_PACKAGE_BASE = 'http://schemas.openxmlformats.org/package/2006/relationships/';

export const RelType = {
  officeDocument: `${REL_BASE}officeDocument`,
  worksheet: `${REL_BASE}worksheet`,
  chartsheet: `${REL_BASE}chartsheet`,
  dialogsheet: `${REL_BASE}dialogsheet`,
  sharedStrings: `${REL_BASE}sharedStrings`,
  styles: `${REL_BASE}styles`,
  theme: `${REL_BASE}theme`,
  hyperlink: `${REL_BASE}hyperlink`,
  image: `${REL_BASE}image`,
  drawing: `${REL_BASE}drawing`,
  comments: `${REL_BASE}comments`,
  vmlDrawing: `${REL_BASE}vmlDrawing`,
  table: `${REL_BASE}table`,
  pivotTable: `${REL_BASE}pivotTable`,
  pivotCacheDefinition: `${REL_BASE}pivotCacheDefinition`,
  pivotCacheRecords: `${REL_BASE}pivotCacheRecords`,
  externalLink: `${REL_BASE}externalLink`,
  externalLinkPath: `${REL_BASE}externalLinkPath`,
  vbaProject: `${REL_BASE}vbaProject`,
  chart: `${REL_BASE}chart`,
  coreProperties: `${REL_PACKAGE_BASE}metadata/core-properties`,
  extendedProperties: `${REL_BASE}extendedProperties`,
  customUI: 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',
  customUI14: 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility',
  threadedComment: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
  person: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
  sheetMetadata: `${REL_BASE}sheetMetadata`,
} as const;

/** Content types that identify the four workbook flavours. */
export const ContentType = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  // Note the vnd.ms-excel namespace on the macro-enabled variants; using the
  // openxmlformats one here is a guaranteed repair prompt.
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
  xlam: 'application/vnd.ms-excel.addin.macroEnabled.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  chartex: 'application/vnd.ms-office.chartex+xml',
  coreProps: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProps: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  vbaProject: 'application/vnd.ms-office.vbaProject',
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
  sheetMetadata: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
} as const;

/** The workbook flavours, keyed by the content type of the workbook part. */
export type WorkbookFlavour = 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'xlam';

const FLAVOUR_BY_CONTENT_TYPE: Record<string, WorkbookFlavour> = {
  [ContentType.xlsx]: 'xlsx',
  [ContentType.xlsm]: 'xlsm',
  [ContentType.xltx]: 'xltx',
  [ContentType.xltm]: 'xltm',
  [ContentType.xlam]: 'xlam',
};

export interface Relationship {
  id: string;
  type: string;
  target: string;
  /** "External" for hyperlinks and linked workbooks, which are not package parts. */
  targetMode?: 'Internal' | 'External';
}

export class OpcError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'OpcError';
  }
}

/**
 * Namespace normalisation for Strict ISO 29500 files.
 *
 * Excel's "Strict Open XML Spreadsheet" save option produces a .xlsx whose every
 * namespace URI is swapped to the purl.oclc.org/ooxml family. A parser that
 * matches transitional URIs literally sees an empty document. We normalise on
 * read and remember the conformance class so a save writes back the same way.
 */
const STRICT_PREFIX = 'http://purl.oclc.org/ooxml/';
const TRANSITIONAL_PREFIX = 'http://schemas.openxmlformats.org/';

export type ConformanceClass = 'transitional' | 'strict';

export function detectConformance(xml: string): ConformanceClass {
  return xml.includes(STRICT_PREFIX) ? 'strict' : 'transitional';
}

/**
 * Rewrite Strict namespace URIs to their Transitional equivalents so the rest of
 * the reader only has to know one set of names.
 */
export function normaliseNamespaces(xml: string): string {
  if (!xml.includes(STRICT_PREFIX)) return xml;
  return xml
    .replaceAll(`${STRICT_PREFIX}spreadsheetml/main`, `${TRANSITIONAL_PREFIX}spreadsheetml/2006/main`)
    .replaceAll(
      `${STRICT_PREFIX}officeDocument/relationships`,
      `${TRANSITIONAL_PREFIX}officeDocument/2006/relationships`,
    )
    .replaceAll(`${STRICT_PREFIX}drawingml/main`, `${TRANSITIONAL_PREFIX}drawingml/2006/main`)
    .replaceAll(`${STRICT_PREFIX}drawingml/chart`, `${TRANSITIONAL_PREFIX}drawingml/2006/chart`)
    .replaceAll(
      `${STRICT_PREFIX}officeDocument/extendedProperties`,
      `${TRANSITIONAL_PREFIX}officeDocument/2006/extended-properties`,
    );
}

/**
 * Reject part names that could escape the package if ever written to disk.
 *
 * OOXML parts are addressed inside the zip, so "zip slip" does not bite us
 * directly - but an export or debug-dump feature that writes parts out would
 * inherit the hole, and it costs nothing to refuse the names here.
 */
function assertSafePartName(name: string): void {
  if (name.startsWith('/') || name.startsWith('\\') || /^[A-Za-z]:/.test(name)) {
    throw new OpcError(`unsafe absolute part name: ${name}`);
  }
  if (name.split(/[\\/]/).includes('..')) {
    throw new OpcError(`unsafe part name escapes the package: ${name}`);
  }
}

/**
 * Every entry in the package, plus the relationship graph.
 *
 * Parts are keyed by their package path without a leading slash, matching how
 * zip entries are named: `xl/worksheets/sheet1.xml`.
 */
export class OpcPackage {
  /** Raw bytes of every entry, in the archive's original order. */
  readonly parts = new Map<string, Uint8Array>();
  /** Archive order, preserved so a re-save keeps the original layout. */
  readonly order: string[] = [];
  /** Content type per part, resolved from Defaults and Overrides. */
  private readonly overrides = new Map<string, string>();
  private readonly defaults = new Map<string, string>();
  /** Relationship tables, keyed by the *source part* they belong to. */
  private readonly rels = new Map<string, Relationship[]>();
  conformance: ConformanceClass = 'transitional';

  static read(bytes: Uint8Array): OpcPackage {
    const pkg = new OpcPackage();
    const entries = readZip(bytes);
    for (const [name, entry] of entries) {
      assertSafePartName(name);
      // Directory entries carry no data and are not parts.
      if (name.endsWith('/')) continue;
      pkg.parts.set(name, entry.data());
      pkg.order.push(name);
    }
    pkg.parseContentTypes(entries);
    pkg.parseAllRelationships();
    return pkg;
  }

  private parseContentTypes(entries: Map<string, ZipEntry>): void {
    const ct = entries.get('[Content_Types].xml');
    if (!ct) throw new OpcError('not an OPC package: [Content_Types].xml is missing');
    const xml = new TextDecoder().decode(ct.data());
    this.conformance = detectConformance(xml);
    const r = new XmlReader(xml);
    for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
      if (t !== XmlToken.Open) continue;
      if (r.localName === 'Default') {
        const ext = r.attr('Extension');
        const type = r.attr('ContentType');
        if (ext && type) this.defaults.set(ext.toLowerCase(), type);
      } else if (r.localName === 'Override') {
        const partName = r.attr('PartName');
        const type = r.attr('ContentType');
        if (partName && type) this.overrides.set(stripLeadingSlash(partName), type);
      }
    }
  }

  private parseAllRelationships(): void {
    for (const name of this.order) {
      if (!name.endsWith('.rels')) continue;
      const source = relsSourcePart(name);
      this.rels.set(source, parseRelationships(this.text(name)));
    }
  }

  /** Content type of a part: an explicit Override, else the extension Default. */
  contentType(part: string): string | undefined {
    const explicit = this.overrides.get(part);
    if (explicit) return explicit;
    const dot = part.lastIndexOf('.');
    return dot < 0 ? undefined : this.defaults.get(part.slice(dot + 1).toLowerCase());
  }

  has(part: string): boolean {
    return this.parts.has(part);
  }

  bytes(part: string): Uint8Array | undefined {
    return this.parts.get(part);
  }

  /** Part decoded as text, with Strict namespaces normalised to Transitional. */
  text(part: string): string {
    const raw = this.parts.get(part);
    if (!raw) throw new OpcError(`part not found: ${part}`);
    return normaliseNamespaces(new TextDecoder().decode(raw));
  }

  /** Relationships declared *by* a part (its `_rels/<name>.rels` sibling). */
  relationshipsOf(part: string): Relationship[] {
    return this.rels.get(part) ?? [];
  }

  /** Resolve one relationship id declared by `part` to a package path. */
  resolve(part: string, id: string): string | undefined {
    const rel = this.relationshipsOf(part).find((r) => r.id === id);
    if (!rel || rel.targetMode === 'External') return undefined;
    return resolveTarget(part, rel.target);
  }

  /** All parts `part` points at with a given relationship type. */
  related(part: string, type: string): { rel: Relationship; path: string }[] {
    const out: { rel: Relationship; path: string }[] = [];
    for (const rel of this.relationshipsOf(part)) {
      if (rel.type !== type) continue;
      if (rel.targetMode === 'External') continue;
      out.push({ rel, path: resolveTarget(part, rel.target) });
    }
    return out;
  }

  /** The package-level relationship to the main document part. */
  mainDocumentPath(): string {
    const root = this.relationshipsOf('');
    const rel = root.find((r) => r.type === RelType.officeDocument);
    if (!rel) throw new OpcError('package has no officeDocument relationship');
    return stripLeadingSlash(rel.target);
  }

  flavour(): WorkbookFlavour {
    const type = this.contentType(this.mainDocumentPath());
    return (type === undefined ? undefined : FLAVOUR_BY_CONTENT_TYPE[type]) ?? 'xlsx';
  }

  /** True when the package carries a VBA project. */
  hasMacros(): boolean {
    return this.parts.has('xl/vbaProject.bin');
  }

  put(part: string, data: Uint8Array, contentType?: string): void {
    assertSafePartName(part);
    if (!this.parts.has(part)) this.order.push(part);
    this.parts.set(part, data);
    if (contentType) this.overrides.set(part, contentType);
  }

  putText(part: string, xml: string, contentType?: string): void {
    this.put(part, new TextEncoder().encode(xml), contentType);
  }

  /**
   * Remove a part along with its content-type Override and any relationship
   * pointing at it. Used for calcChain.xml, which is the one part we
   * deliberately destroy rather than preserve (see `dropCalcChain`).
   */
  remove(part: string): void {
    this.parts.delete(part);
    const i = this.order.indexOf(part);
    if (i >= 0) this.order.splice(i, 1);
    this.overrides.delete(part);
    for (const [source, list] of this.rels) {
      const kept = list.filter((r) => resolveTarget(source, r.target) !== part);
      if (kept.length !== list.length) {
        this.rels.set(source, kept);
        this.writeRelsPart(source, kept);
      }
    }
  }

  /**
   * Delete `xl/calcChain.xml`.
   *
   * The calculation chain is a cache of the order Excel last evaluated formulas
   * in. Once we have edited any formula it is stale, and a stale chain makes
   * Excel show a repair prompt naming the part. Updating it correctly is
   * possible but pointless: the supported fix is to delete it and let Excel
   * rebuild, which is why this is the single deliberate exception to the
   * otherwise absolute rule that we preserve what we do not understand.
   */
  dropCalcChain(): void {
    if (this.parts.has('xl/calcChain.xml')) this.remove('xl/calcChain.xml');
  }

  private writeRelsPart(source: string, list: Relationship[]): void {
    const path = relsPathFor(source);
    const w = new XmlWriter();
    w.open('Relationships', { xmlns: `${REL_PACKAGE_BASE.slice(0, -14)}relationships` });
    for (const rel of list) {
      w.empty('Relationship', {
        Id: rel.id,
        Type: rel.type,
        Target: rel.target,
        TargetMode: rel.targetMode === 'External' ? 'External' : undefined,
      });
    }
    w.close();
    this.putText(path, w.toString());
  }

  /**
   * Mint a relationship id that does not collide with any existing one.
   * Existing ids are never renumbered: `r:id="rId3"` is referenced from inside
   * the part's own XML, and "tidying" the numbering silently breaks hyperlinks,
   * drawings, comments, tables, and external links.
   */
  addRelationship(source: string, type: string, target: string, external = false): string {
    const list = this.rels.get(source) ?? [];
    let n = list.length + 1;
    const used = new Set(list.map((r) => r.id));
    while (used.has(`rId${n}`)) n++;
    const rel: Relationship = {
      id: `rId${n}`,
      type,
      target,
      targetMode: external ? 'External' : 'Internal',
    };
    list.push(rel);
    this.rels.set(source, list);
    this.writeRelsPart(source, list);
    return rel.id;
  }

  /** Rebuild `[Content_Types].xml` from the current Defaults and Overrides. */
  private renderContentTypes(): string {
    const w = new XmlWriter();
    w.open('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' });
    const defaults = new Map(this.defaults);
    defaults.set('rels', ContentType.relationships);
    defaults.set('xml', ContentType.xml);
    // A Default is required for every binary extension actually present, or
    // Excel cannot type the part at all.
    for (const part of this.order) {
      const dot = part.lastIndexOf('.');
      if (dot < 0) continue;
      const ext = part.slice(dot + 1).toLowerCase();
      if (defaults.has(ext) || this.overrides.has(part)) continue;
      const guessed = DEFAULT_BY_EXTENSION[ext];
      if (guessed) defaults.set(ext, guessed);
    }
    for (const [ext, type] of [...defaults].sort((a, b) => a[0].localeCompare(b[0]))) {
      w.empty('Default', { Extension: ext, ContentType: type });
    }
    for (const part of this.order) {
      const type = this.overrides.get(part);
      if (type) w.empty('Override', { PartName: `/${part}`, ContentType: type });
    }
    w.close();
    return w.toString();
  }

  /** Serialise back to a zip. */
  write(options: { modified?: Date } = {}): Uint8Array {
    this.putText('[Content_Types].xml', this.renderContentTypes());

    const entries: ZipWriteEntry[] = [];
    // Convention (and some third-party readers' requirement): content types first.
    entries.push({ name: '[Content_Types].xml', data: this.parts.get('[Content_Types].xml')! });
    for (const name of this.order) {
      if (name === '[Content_Types].xml') continue;
      const data = this.parts.get(name);
      if (!data) continue;
      entries.push({ name, data, store: shouldStore(name) });
    }
    return writeZip(entries, { modified: options.modified });
  }
}

/** Binary formats that are already compressed; deflating them again just costs time. */
function shouldStore(name: string): boolean {
  return /\.(png|jpe?g|gif|zip|gz|mp4|webp)$/i.test(name);
}

const DEFAULT_BY_EXTENSION: Record<string, string> = {
  bin: ContentType.vbaProject,
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export function parseRelationships(xml: string): Relationship[] {
  const out: Relationship[] = [];
  const r = new XmlReader(xml);
  for (let t = r.next(); t !== XmlToken.EOF; t = r.next()) {
    if (t !== XmlToken.Open || r.localName !== 'Relationship') continue;
    const id = r.attr('Id');
    const type = r.attr('Type');
    const target = r.attr('Target');
    if (!id || !type || target === undefined) continue;
    const mode = r.attr('TargetMode');
    out.push({
      id,
      type,
      target,
      targetMode: mode === 'External' ? 'External' : 'Internal',
    });
  }
  return out;
}

function stripLeadingSlash(s: string): string {
  return s.startsWith('/') ? s.slice(1) : s;
}

/** `xl/_rels/workbook.xml.rels` -> `xl/workbook.xml`; `_rels/.rels` -> ``. */
export function relsSourcePart(relsPath: string): string {
  const i = relsPath.lastIndexOf('_rels/');
  if (i < 0) return '';
  const dir = relsPath.slice(0, i);
  const file = relsPath.slice(i + 6).replace(/\.rels$/, '');
  return file === '' ? '' : dir + file;
}

/** Inverse of `relsSourcePart`. */
export function relsPathFor(part: string): string {
  if (part === '') return '_rels/.rels';
  const slash = part.lastIndexOf('/');
  const dir = slash < 0 ? '' : part.slice(0, slash + 1);
  const file = slash < 0 ? part : part.slice(slash + 1);
  return `${dir}_rels/${file}.rels`;
}

/**
 * Resolve a relationship target against the part that declared it.
 *
 * Targets are relative to the *directory* of the source part, except when they
 * start with a slash, which makes them package-absolute.
 */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return stripLeadingSlash(target);
  const slash = sourcePart.lastIndexOf('/');
  const dir = slash < 0 ? '' : sourcePart.slice(0, slash + 1);
  const segments = (dir + target).split('/');
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === '.' || seg === '') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}
