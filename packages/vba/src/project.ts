/**
 * VBA project structure: reading vbaProject.bin, per [MS-OVBA] section 2.3.
 *
 * A macro-enabled workbook carries its code in a compound file inside the
 * package. That file holds a plain-text PROJECT stream listing the modules, a
 * compressed 'VBA/dir' stream describing them in binary, and one stream per
 * module whose second half is the compressed source text. This module walks all
 * three and hands back the source.
 *
 * SECURITY. Nothing here executes anything. This is a reader: it turns a
 * container into strings, and the strings are destined for a read-only viewer.
 * Extraction is safe in a way that execution is not; when we come to run macros
 * that will need an explicit sandbox and explicit consent from the user, and
 * none of that belongs in a parser.
 *
 * PROTECTION. A project can be locked for viewing, and we surface that as a
 * flag and then stop. We deliberately do not implement removal or bypass of it.
 * The obfuscation over the CMG/DPB/GC values is trivially reversible and
 * documented in section 2.4.3, so "we could" is not the question; the answer is
 * that a spreadsheet application has no business stripping the protection its
 * user's colleague asked for. We decode exactly enough of those blobs to answer
 * "is this locked" - the protection flags, and the *length* of the password
 * field - and never touch the password bytes themselves.
 *
 * The parser is written to survive files it does not understand. Real projects
 * come from twenty years of Office versions and carry records this code has
 * never heard of, so an unknown record is skipped by its declared size and
 * noted in `warnings` rather than aborting the parse. Losing one module's
 * description is a much better outcome than showing the user nothing.
 */

import { type CfbFile, readCfb } from '@mirrorz/formats';
import { VbaCompressionError, decompress } from './compression.js';

/** Raised when the container is not a VBA project at all. */
export class VbaProjectError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'VbaProjectError';
  }
}

// --- dir stream record identifiers, section 2.3.4.2 --------------------------

const ID_PROJECTSYSKIND = 0x0001;
const ID_PROJECTLCID = 0x0002;
const ID_PROJECTCODEPAGE = 0x0003;
const ID_PROJECTNAME = 0x0004;
const ID_PROJECTDOCSTRING = 0x0005;
const ID_PROJECTHELPFILEPATH = 0x0006;
const ID_PROJECTHELPCONTEXT = 0x0007;
const ID_PROJECTLIBFLAGS = 0x0008;
const ID_PROJECTVERSION = 0x0009;
const ID_PROJECTCONSTANTS = 0x000c;
const ID_REFERENCEREGISTERED = 0x000d;
const ID_REFERENCEPROJECT = 0x000e;
const ID_PROJECTMODULES = 0x000f;
const ID_DIR_TERMINATOR = 0x0010;
const ID_PROJECTCOOKIE = 0x0013;
const ID_PROJECTLCIDINVOKE = 0x0014;
const ID_REFERENCENAME = 0x0016;
const ID_MODULENAME = 0x0019;
const ID_MODULESTREAMNAME = 0x001a;
const ID_MODULEDOCSTRING = 0x001c;
const ID_MODULEHELPCONTEXT = 0x001e;
const ID_MODULETYPE_PROCEDURAL = 0x0021;
const ID_MODULETYPE_DOCUMENT = 0x0022;
const ID_MODULEREADONLY = 0x0025;
const ID_MODULEPRIVATE = 0x0028;
const ID_MODULE_TERMINATOR = 0x002b;
const ID_MODULECOOKIE = 0x002c;
const ID_REFERENCECONTROL = 0x002f;
const ID_REFERENCECONTROL_EXTENDED = 0x0030;
const ID_MODULEOFFSET = 0x0031;
const ID_REFERENCEORIGINAL = 0x0033;
const ID_MODULENAMEUNICODE = 0x0047;
const ID_PROJECTCOMPATVERSION = 0x004a;

// --- public shape ------------------------------------------------------------

/**
 * Procedural modules hold ordinary Sub/Function code. Everything else - a
 * worksheet's code-behind, a class module, a UserForm's designer - shares one
 * record identifier; the PROJECT stream is what tells those three apart, and it
 * is exposed as `properties` for callers that care.
 */
export type VbaModuleType = 'procedural' | 'document';

export interface VbaModule {
  /** Name as VBA shows it in the project explorer. */
  name: string;
  /** Name of the stream inside the VBA storage holding the code. */
  streamName: string;
  type: VbaModuleType;
  /** Decoded source text, empty when the stream was missing or unreadable. */
  source: string;
  /** MODULEOFFSET: where the compressed source begins in the module stream. */
  offset: number;
  /** The module's description, as typed into the object browser. */
  docString: string;
  helpContext: number;
  cookie: number;
  readOnly: boolean;
  /** Private modules are not exposed to other projects referencing this one. */
  isPrivate: boolean;
}

export type VbaReferenceKind = 'registered' | 'project' | 'original' | 'control';

export interface VbaReference {
  kind: VbaReferenceKind;
  /** REFERENCENAME preceding the reference record, when one was present. */
  name: string;
  /** The primary library identifier, whichever field carries it for this kind. */
  libid: string;
  /** REFERENCEPROJECT only: the path-relative form of the identifier. */
  libidRelative?: string;
  majorVersion?: number;
  minorVersion?: number;
}

/** Well-known PROJECT stream properties, section 2.3.1. */
export interface VbaProjectProperties {
  /** ID=, the CLSID of the project's type library. */
  id: string;
  name: string;
  helpFile: string;
  helpContextId: string;
  description: string;
  versionCompatible32: string;
  /** Document=, one per worksheet or workbook code-behind module. */
  documents: string[];
  /** Module=, the procedural modules. */
  modules: string[];
  /** Class=, the class modules. */
  classes: string[];
  /** BaseClass=, the designer modules such as UserForms. */
  baseClasses: string[];
  /** Every key=value line, in file order, including the ones above. */
  entries: Array<[string, string]>;
  /** The stream as text, so a caller can show what we did not model. */
  text: string;
}

export interface VbaProtection {
  /** The user ticked "Lock project for viewing". */
  userProtected: boolean;
  /** The host application asked for the project to be protected. */
  hostProtected: boolean;
  /** The VBA editor asked for the project to be protected. */
  vbeProtected: boolean;
  /**
   * A project password is set. Derived from the length of the DPB payload, not
   * from its contents: the password itself is none of our business.
   */
  passwordSet: boolean;
}

export interface VbaProject {
  projectName: string;
  /** PROJECTCODEPAGE, the code page every MBCS string in the project uses. */
  codePage: number;
  /**
   * The project is locked or password protected. When true, `modules` may still
   * list what is in the project but the source is not ours to show.
   */
  protected: boolean;
  protection: VbaProtection;
  modules: VbaModule[];
  references: VbaReference[];
  properties: VbaProjectProperties;
  /** PROJECTSYSKIND: 0 Win16, 1 Win32, 2 Macintosh, 3 Win64. */
  sysKind: number;
  lcid: number;
  docString: string;
  helpFile: string;
  /** PROJECTCONSTANTS, the conditional compilation constants. */
  constants: string;
  versionMajor: number;
  versionMinor: number;
  /** Everything we could not make sense of, in the order we met it. */
  warnings: string[];
}

export interface ParseVbaProjectOptions {
  /**
   * Extract module source even when the project is protected. Off by default:
   * protection is a request from whoever wrote the file, and the default should
   * honour it. Setting this decodes nothing that is encrypted - VBA source is
   * merely compressed, never encrypted - it only declines to stop.
   */
  includeProtectedSource?: boolean;
}

// --- code pages --------------------------------------------------------------

/**
 * The 0x80-0x9F range is what separates windows-1252 from ISO-8859-1, and it is
 * exactly the range that carries the curly quotes and dashes people paste into
 * comments. Spelled out rather than fetched from ICU so that the result does not
 * depend on how the host's Node was built.
 */
const CP1252_HIGH = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039,
  0x0152, 0x008d, 0x017d, 0x008f, 0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
];

/** Code pages we can name for the platform decoder; anything else falls back. */
const CODE_PAGE_LABELS = new Map<number, string>([
  [874, 'windows-874'],
  [932, 'shift_jis'],
  [936, 'gbk'],
  [949, 'euc-kr'],
  [950, 'big5'],
  [1250, 'windows-1250'],
  [1251, 'windows-1251'],
  [1253, 'windows-1253'],
  [1254, 'windows-1254'],
  [1255, 'windows-1255'],
  [1256, 'windows-1256'],
  [1257, 'windows-1257'],
  [1258, 'windows-1258'],
  [10000, 'macintosh'],
  [20866, 'koi8-r'],
  [21866, 'koi8-u'],
  [28591, 'iso-8859-1'],
  [28592, 'iso-8859-2'],
  [28605, 'iso-8859-15'],
]);

function decodeCp1252(bytes: Uint8Array): string {
  let text = '';
  for (const byte of bytes) {
    text += String.fromCharCode(byte >= 0x80 && byte <= 0x9f ? CP1252_HIGH[byte - 0x80]! : byte);
  }
  return text;
}

/**
 * Decode MBCS bytes in the project's code page.
 *
 * 1252 and 65001 are implemented outright because between them they cover
 * essentially every file a western user opens. For the rest we ask the platform
 * decoder, which a full-ICU Node can satisfy, and if it cannot we fall back to
 * windows-1252 and say so in the warnings. The fallback is deliberate: it is
 * byte-preserving for ASCII, which is what VBA keywords and identifiers are, so
 * the code stays readable even when a comment in Cyrillic does not.
 */
function decodeText(bytes: Uint8Array, codePage: number, warnings: string[]): string {
  // The same complaint would otherwise arrive once per string in the project.
  const warnOnce = (message: string): void => {
    if (!warnings.includes(message)) warnings.push(message);
  };
  if (bytes.length === 0) return '';
  if (codePage === 1252 || codePage === 0) return decodeCp1252(bytes);
  if (codePage === 65001) return new TextDecoder('utf-8').decode(bytes);
  const label = CODE_PAGE_LABELS.get(codePage);
  if (label !== undefined) {
    try {
      return new TextDecoder(label).decode(bytes);
    } catch {
      warnOnce(`code page ${codePage} is unavailable on this platform; decoded as windows-1252`);
      return decodeCp1252(bytes);
    }
  }
  warnOnce(`code page ${codePage} is not implemented; decoded as windows-1252`);
  return decodeCp1252(bytes);
}

// --- a bounds-checked cursor -------------------------------------------------

class Cursor {
  pos = 0;
  constructor(readonly data: Uint8Array) {}

  get remaining(): number {
    return this.data.length - this.pos;
  }

  u16(): number {
    if (this.remaining < 2) throw new VbaProjectError(`truncated at ${this.pos}: wanted 2 bytes`);
    const value = this.data[this.pos]! | (this.data[this.pos + 1]! << 8);
    this.pos += 2;
    return value;
  }

  u32(): number {
    if (this.remaining < 4) throw new VbaProjectError(`truncated at ${this.pos}: wanted 4 bytes`);
    const value =
      (this.data[this.pos]! |
        (this.data[this.pos + 1]! << 8) |
        (this.data[this.pos + 2]! << 16) |
        (this.data[this.pos + 3]! << 24)) >>>
      0;
    this.pos += 4;
    return value;
  }

  bytes(length: number): Uint8Array {
    if (length > this.remaining) {
      throw new VbaProjectError(`truncated at ${this.pos}: wanted ${length} bytes`);
    }
    const slice = this.data.subarray(this.pos, this.pos + length);
    this.pos += length;
    return slice;
  }

  skip(length: number): void {
    if (length > this.remaining) {
      throw new VbaProjectError(`truncated at ${this.pos}: wanted ${length} bytes`);
    }
    this.pos += length;
  }

  /** Look at the next record identifier without consuming it. */
  peek16(): number | undefined {
    if (this.remaining < 2) return undefined;
    return this.data[this.pos]! | (this.data[this.pos + 1]! << 8);
  }
}

// --- dir stream --------------------------------------------------------------

interface DirModule {
  name: string;
  nameUnicode: string;
  streamName: string;
  type: VbaModuleType;
  offset: number;
  docString: string;
  helpContext: number;
  cookie: number;
  readOnly: boolean;
  isPrivate: boolean;
}

/** A reference's records as read, before the code page is known. */
interface RawReference {
  kind: VbaReferenceKind;
  name?: Uint8Array;
  libid?: Uint8Array;
  libidRelative?: Uint8Array;
  majorVersion?: number;
  minorVersion?: number;
}

/**
 * A module's records as read. Text stays as bytes until the whole stream has
 * been walked, because PROJECTCODEPAGE is not guaranteed to have been seen by
 * the time the first string turns up.
 */
interface RawModule {
  name?: Uint8Array;
  nameUnicode?: Uint8Array;
  streamName?: Uint8Array;
  docString?: Uint8Array;
  type: VbaModuleType;
  offset: number;
  helpContext: number;
  cookie: number;
  readOnly: boolean;
  isPrivate: boolean;
}

interface DirStream {
  sysKind: number;
  lcid: number;
  codePage: number;
  name: string;
  docString: string;
  helpFile: string;
  constants: string;
  versionMajor: number;
  versionMinor: number;
  references: VbaReference[];
  modules: DirModule[];
}

/**
 * Most dir records are an identifier, a four-byte size and that many bytes of
 * payload, which is what makes skipping the unknown ones safe.
 */
function readSizedBytes(cursor: Cursor): Uint8Array {
  return cursor.bytes(cursor.u32());
}

/**
 * Several records store the same string twice, once as MBCS and once as UTF-16,
 * separated by a two-byte reserved field. Only the MBCS half is returned; the
 * unicode half is a copy and Office does not always keep it in step.
 */
function readMbcsPair(cursor: Cursor): Uint8Array {
  const mbcs = readSizedBytes(cursor);
  cursor.u16();
  readSizedBytes(cursor);
  return mbcs;
}

/**
 * Parse the decompressed dir stream. Anything unrecognised is skipped by its
 * declared size; anything that leaves us unable to find the next record ends the
 * walk with a warning rather than an exception, so a partially understood
 * project still yields its modules.
 */
function parseDirStream(data: Uint8Array, warnings: string[]): DirStream {
  const cursor = new Cursor(data);
  // Strings appear before PROJECTCODEPAGE is guaranteed to have been seen, so
  // the raw bytes are held and decoded at the end once the code page is known.
  const raw = new Map<number, Uint8Array>();
  const result: DirStream = {
    sysKind: 0,
    lcid: 0,
    codePage: 1252,
    name: '',
    docString: '',
    helpFile: '',
    constants: '',
    versionMajor: 0,
    versionMinor: 0,
    references: [],
    modules: [],
  };
  const rawReferences: RawReference[] = [];
  const rawModules: RawModule[] = [];

  let pendingReferenceName: Uint8Array | undefined;
  let moduleCount = 0;
  let sawModulesRecord = false;

  try {
    walk: while (cursor.remaining >= 2) {
      const id = cursor.u16();
      switch (id) {
        case ID_PROJECTSYSKIND:
          cursor.u32();
          result.sysKind = cursor.u32();
          break;
        case ID_PROJECTLCID:
          cursor.u32();
          result.lcid = cursor.u32();
          break;
        case ID_PROJECTLCIDINVOKE:
        case ID_PROJECTHELPCONTEXT:
        case ID_PROJECTLIBFLAGS:
        case ID_PROJECTCOMPATVERSION:
          cursor.skip(cursor.u32());
          break;
        case ID_PROJECTCODEPAGE:
          cursor.u32();
          result.codePage = cursor.u16();
          break;
        case ID_PROJECTNAME:
          raw.set(ID_PROJECTNAME, readSizedBytes(cursor));
          break;
        case ID_PROJECTDOCSTRING:
          raw.set(ID_PROJECTDOCSTRING, readMbcsPair(cursor));
          break;
        case ID_PROJECTHELPFILEPATH:
          raw.set(ID_PROJECTHELPFILEPATH, readMbcsPair(cursor));
          break;
        case ID_PROJECTCONSTANTS:
          raw.set(ID_PROJECTCONSTANTS, readMbcsPair(cursor));
          break;
        case ID_PROJECTVERSION:
          // The odd one out: its size field is a reserved 0x00000004 that does
          // not describe the six bytes of payload that follow it.
          cursor.u32();
          result.versionMajor = cursor.u32();
          result.versionMinor = cursor.u16();
          break;
        case ID_REFERENCENAME:
          pendingReferenceName = readMbcsPair(cursor);
          break;
        case ID_REFERENCEREGISTERED: {
          const body = new Cursor(readSizedBytes(cursor));
          rawReferences.push({
            kind: 'registered',
            name: pendingReferenceName,
            libid: readSizedBytes(body),
          });
          pendingReferenceName = undefined;
          break;
        }
        case ID_REFERENCEPROJECT: {
          const body = new Cursor(readSizedBytes(cursor));
          const libid = readSizedBytes(body);
          const libidRelative = readSizedBytes(body);
          rawReferences.push({
            kind: 'project',
            name: pendingReferenceName,
            libid,
            libidRelative,
            majorVersion: body.remaining >= 4 ? body.u32() : 0,
            minorVersion: body.remaining >= 2 ? body.u16() : 0,
          });
          pendingReferenceName = undefined;
          break;
        }
        case ID_REFERENCEORIGINAL:
          // Always followed by the REFERENCECONTROL it describes, which the next
          // turn of this loop picks up, so the pending name carries over to it.
          rawReferences.push({
            kind: 'original',
            name: pendingReferenceName,
            libid: readSizedBytes(cursor),
          });
          break;
        case ID_REFERENCECONTROL: {
          const body = new Cursor(readSizedBytes(cursor));
          rawReferences.push({
            kind: 'control',
            name: pendingReferenceName,
            libid: readSizedBytes(body),
          });
          pendingReferenceName = undefined;
          // The extended half is optional and may be preceded by its own name
          // record, so it is identified by peeking rather than by position.
          if (cursor.peek16() === ID_REFERENCENAME) {
            cursor.u16();
            readMbcsPair(cursor);
          }
          if (cursor.peek16() === ID_REFERENCECONTROL_EXTENDED) {
            cursor.u16();
            cursor.skip(cursor.u32());
          }
          break;
        }
        case ID_PROJECTMODULES: {
          cursor.u32();
          moduleCount = cursor.u16();
          if (cursor.peek16() === ID_PROJECTCOOKIE) {
            cursor.u16();
            cursor.skip(cursor.u32());
          }
          sawModulesRecord = true;
          break walk;
        }
        case ID_DIR_TERMINATOR:
          // The terminator is Id plus a four byte reserved field. Consuming
          // only the identifier would leave those four bytes to be read as the
          // start of a module, which produces two invented warnings on a file
          // that is in fact perfectly well formed.
          cursor.skip(Math.min(4, cursor.remaining));
          break walk;
        default:
          warnings.push(`dir: skipped unknown record 0x${id.toString(16).padStart(4, '0')}`);
          cursor.skip(cursor.u32());
          break;
      }
    }

    // Modules run until the dir terminator; the declared count is treated as a
    // sanity bound rather than gospel because a truncated file will lie. A
    // stream that ended at the terminator without a PROJECTMODULES record has
    // no module array at all, and whatever follows is not one.
    modules: while (sawModulesRecord && cursor.remaining >= 2) {
      const first = cursor.peek16();
      if (first === ID_DIR_TERMINATOR || first === undefined) break;
      const module: RawModule = {
        type: 'procedural',
        offset: 0,
        helpContext: 0,
        cookie: 0,
        readOnly: false,
        isPrivate: false,
      };
      for (;;) {
        if (cursor.remaining < 2) break modules;
        const id = cursor.u16();
        if (id === ID_MODULE_TERMINATOR) {
          cursor.u32();
          break;
        }
        switch (id) {
          case ID_MODULENAME:
            module.name = readSizedBytes(cursor);
            break;
          case ID_MODULENAMEUNICODE:
            module.nameUnicode = readSizedBytes(cursor);
            break;
          case ID_MODULESTREAMNAME:
            module.streamName = readMbcsPair(cursor);
            break;
          case ID_MODULEDOCSTRING:
            module.docString = readMbcsPair(cursor);
            break;
          case ID_MODULEOFFSET:
            cursor.u32();
            module.offset = cursor.u32();
            break;
          case ID_MODULEHELPCONTEXT:
            cursor.u32();
            module.helpContext = cursor.u32();
            break;
          case ID_MODULECOOKIE:
            cursor.u32();
            module.cookie = cursor.u16();
            break;
          case ID_MODULETYPE_PROCEDURAL:
            cursor.u32();
            module.type = 'procedural';
            break;
          case ID_MODULETYPE_DOCUMENT:
            cursor.u32();
            module.type = 'document';
            break;
          case ID_MODULEREADONLY:
            cursor.u32();
            module.readOnly = true;
            break;
          case ID_MODULEPRIVATE:
            cursor.u32();
            module.isPrivate = true;
            break;
          default:
            warnings.push(`dir: skipped unknown module record 0x${id.toString(16).padStart(4, '0')}`);
            cursor.skip(cursor.u32());
            break;
        }
      }
      rawModules.push(module);
      if (rawModules.length > moduleCount && moduleCount > 0) {
        warnings.push(`dir: more modules than the declared count of ${moduleCount}`);
        break;
      }
    }
  } catch (error) {
    warnings.push(`dir: stopped early (${(error as Error).message})`);
  }

  if (moduleCount > 0 && rawModules.length < moduleCount) {
    warnings.push(`dir: declared ${moduleCount} modules but found ${rawModules.length}`);
  }

  const text = (bytes: Uint8Array | undefined): string =>
    bytes === undefined ? '' : decodeText(bytes, result.codePage, warnings);

  result.name = text(raw.get(ID_PROJECTNAME));
  result.docString = text(raw.get(ID_PROJECTDOCSTRING));
  result.helpFile = text(raw.get(ID_PROJECTHELPFILEPATH));
  result.constants = text(raw.get(ID_PROJECTCONSTANTS));
  result.references = rawReferences.map((reference) => {
    const mapped: VbaReference = {
      kind: reference.kind,
      name: text(reference.name),
      libid: text(reference.libid),
    };
    if (reference.libidRelative !== undefined) mapped.libidRelative = text(reference.libidRelative);
    if (reference.majorVersion !== undefined) mapped.majorVersion = reference.majorVersion;
    if (reference.minorVersion !== undefined) mapped.minorVersion = reference.minorVersion;
    return mapped;
  });
  result.modules = rawModules.map((module) => {
    const name = text(module.name);
    const streamName = text(module.streamName);
    return {
      name,
      // MODULENAMEUNICODE is UTF-16LE regardless of the project code page.
      nameUnicode: decodeUtf16(module.nameUnicode),
      streamName: streamName === '' ? name : streamName,
      type: module.type,
      offset: module.offset,
      docString: text(module.docString),
      helpContext: module.helpContext,
      cookie: module.cookie,
      readOnly: module.readOnly,
      isPrivate: module.isPrivate,
    };
  });

  return result;
}

function decodeUtf16(bytes: Uint8Array | undefined): string {
  if (bytes === undefined || bytes.length < 2) return '';
  let text = '';
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    text += String.fromCharCode(bytes[i]! | (bytes[i + 1]! << 8));
  }
  return text;
}

// --- PROJECT stream ----------------------------------------------------------

/**
 * Parse the PROJECT stream, section 2.3.1. It is an INI-like list of key=value
 * lines followed by [Host Extender Info] and [Workspace] sections, and only the
 * lines before the first section header describe the project itself.
 */
function parseProjectStream(text: string): VbaProjectProperties {
  const properties: VbaProjectProperties = {
    id: '',
    name: '',
    helpFile: '',
    helpContextId: '',
    description: '',
    versionCompatible32: '',
    documents: [],
    modules: [],
    classes: [],
    baseClasses: [],
    entries: [],
    text,
  };

  for (const line of text.split(/\r\n|\n|\r/)) {
    const trimmed = line.trim();
    if (trimmed === '') continue;
    if (trimmed.startsWith('[')) break;
    const eq = trimmed.indexOf('=');
    if (eq < 0) continue;
    const key = trimmed.slice(0, eq);
    let value = trimmed.slice(eq + 1);
    if (value.length >= 2 && value.startsWith('"') && value.endsWith('"')) {
      value = value.slice(1, -1);
    }
    properties.entries.push([key, value]);
    switch (key) {
      case 'ID':
        properties.id = value;
        break;
      case 'Name':
        properties.name = value;
        break;
      case 'HelpFile':
        properties.helpFile = value;
        break;
      case 'HelpContextID':
        properties.helpContextId = value;
        break;
      case 'Description':
        properties.description = value;
        break;
      case 'VersionCompatible32':
        properties.versionCompatible32 = value;
        break;
      case 'Document':
        // "Document=Sheet1/&H00000000": the tail is the module's cookie.
        properties.documents.push(value.split('/')[0]!);
        break;
      case 'Module':
        properties.modules.push(value);
        break;
      case 'Class':
        properties.classes.push(value);
        break;
      case 'BaseClass':
        properties.baseClasses.push(value);
        break;
      default:
        break;
    }
  }

  return properties;
}

// --- protection --------------------------------------------------------------

function hexToBytes(hex: string): Uint8Array | undefined {
  if (hex.length === 0 || hex.length % 2 !== 0 || !/^[0-9a-fA-F]+$/.test(hex)) return undefined;
  const bytes = new Uint8Array(hex.length / 2);
  for (let i = 0; i < bytes.length; i++) bytes[i] = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  return bytes;
}

/**
 * Reverse the obfuscation of section 2.4.3 far enough to read the header and
 * the payload length, and no further unless asked.
 *
 * `wantData` is passed for the protection flags, which are the whole point of
 * looking. It is never passed for the password field: reading its length tells
 * us whether one is set, and that is the entire question we are entitled to
 * ask. There is deliberately no encoder here, because writing one of these back
 * is what a protection-removal tool does.
 */
function decodeObfuscated(
  blob: Uint8Array,
  wantData: boolean,
): { version: number; dataLength: number; data?: Uint8Array } | undefined {
  if (blob.length < 3) return undefined;
  const seed = blob[0]!;
  const versionEnc = blob[1]!;
  const projKeyEnc = blob[2]!;
  const projKey = seed ^ projKeyEnc;

  let unencryptedByte1 = projKey;
  let encryptedByte1 = projKeyEnc;
  let encryptedByte2 = versionEnc;
  let pos = 3;

  const next = (): number | undefined => {
    if (pos >= blob.length) return undefined;
    const byteEnc = blob[pos++]!;
    const plain = (byteEnc ^ (encryptedByte2 + unencryptedByte1)) & 0xff;
    encryptedByte2 = encryptedByte1;
    encryptedByte1 = byteEnc;
    unencryptedByte1 = plain;
    return plain;
  };

  // IgnoredEnc exists purely to vary the ciphertext; its length is in the seed.
  const ignoredLength = (seed & 6) / 2;
  for (let i = 0; i < ignoredLength; i++) if (next() === undefined) return undefined;

  let dataLength = 0;
  for (let i = 0; i < 4; i++) {
    const byte = next();
    if (byte === undefined) return undefined;
    dataLength |= byte << (i * 8);
  }
  dataLength >>>= 0;

  const decoded: { version: number; dataLength: number; data?: Uint8Array } = {
    version: seed ^ versionEnc,
    dataLength,
  };
  if (wantData) {
    const available = Math.min(dataLength, blob.length - pos);
    const data = new Uint8Array(available);
    for (let i = 0; i < available; i++) {
      const byte = next();
      if (byte === undefined) break;
      data[i] = byte;
    }
    decoded.data = data;
  }
  return decoded;
}

function readProtection(properties: VbaProjectProperties, warnings: string[]): VbaProtection {
  const protection: VbaProtection = {
    userProtected: false,
    hostProtected: false,
    vbeProtected: false,
    passwordSet: false,
  };

  // A blob we cannot read is reported. Failing quietly here fails open: an
  // unreadable CMG would otherwise mean "not protected", and the source of a
  // project whose author asked for it to be locked would be handed over on the
  // strength of a field we could not parse.
  const decode = (key: string, wantData: boolean) => {
    const value = properties.entries.find(([name]) => name === key)?.[1];
    if (value === undefined) return undefined;
    const blob = hexToBytes(value);
    const decoded = blob === undefined ? undefined : decodeObfuscated(blob, wantData);
    if (decoded === undefined) {
      warnings.push(`${key} is not a readable encrypted structure; protection state unknown`);
      return undefined;
    }
    // Section 2.4.3.2: Version MUST be 2. Anything else means we have not in
    // fact decrypted this, and its bits are not protection flags.
    if (decoded.version !== 2) {
      warnings.push(`${key} decrypted to version ${decoded.version}, expected 2`);
      return undefined;
    }
    return decoded;
  };

  const cmg = decode('CMG', true);
  const data = cmg?.data;
  if (cmg !== undefined && (data === undefined || data.length < 1)) {
    warnings.push('CMG carries no protection state word');
  }
  if (data !== undefined && data.length >= 1) {
    // ProjectProtectionState, section 2.3.1.15: a little-endian 32-bit word
    // whose three lowest bits are fUserProtected, fHostProtected, fVBEProtected.
    const flags = data[0]!;
    protection.userProtected = (flags & 0x01) !== 0;
    protection.hostProtected = (flags & 0x02) !== 0;
    protection.vbeProtected = (flags & 0x04) !== 0;
  }

  // An unprotected project still writes a DPB: one byte of 0x00. Anything
  // longer is a password, either in the clear or hashed, and we do not look.
  const dpb = decode('DPB', false);
  if (dpb !== undefined && dpb.dataLength > 1) protection.passwordSet = true;

  return protection;
}

// --- top level ---------------------------------------------------------------

/**
 * Find a stream by name inside the VBA storage. Most projects put it at
 * 'VBA/<name>', but the storage is not required to be called that in every host,
 * so a bare name and a suffix match are tried before giving up.
 */
function findStream(cfb: CfbFile, streamName: string): Uint8Array | undefined {
  if (cfb.has(`VBA/${streamName}`)) return cfb.read(`VBA/${streamName}`);
  if (cfb.has(streamName)) return cfb.read(streamName);
  const suffix = `/${streamName}`;
  for (const [path, entry] of cfb.entries) {
    if (entry.type === 'stream' && path.endsWith(suffix)) return entry.data();
  }
  return undefined;
}

export function parseVbaProject(bytes: Uint8Array, options: ParseVbaProjectOptions = {}): VbaProject {
  const cfb = readCfb(bytes);
  const warnings: string[] = [];

  const dirBytes = findStream(cfb, 'dir');
  if (dirBytes === undefined) {
    throw new VbaProjectError("no 'VBA/dir' stream: this is not a VBA project");
  }

  let dirData: Uint8Array;
  try {
    dirData = decompress(dirBytes);
  } catch (error) {
    if (error instanceof VbaCompressionError) {
      throw new VbaProjectError(`the dir stream is not decompressible (${error.message})`);
    }
    throw error;
  }
  const dir = parseDirStream(dirData, warnings);

  // PROJECT is MBCS in the project's own code page, which only the dir stream
  // knows, so it has to be decoded second even though it is the plainer of the two.
  const projectBytes = cfb.has('PROJECT') ? cfb.read('PROJECT') : undefined;
  if (projectBytes === undefined) warnings.push("no 'PROJECT' stream");
  const properties = parseProjectStream(
    projectBytes === undefined ? '' : decodeText(projectBytes, dir.codePage, warnings),
  );

  const protection = readProtection(properties, warnings);
  const isProtected =
    protection.userProtected ||
    protection.hostProtected ||
    protection.vbeProtected ||
    protection.passwordSet;

  const modules: VbaModule[] = dir.modules.map((module) => ({
    name: module.name === '' ? module.nameUnicode : module.name,
    streamName: module.streamName,
    type: module.type,
    source: '',
    offset: module.offset,
    docString: module.docString,
    helpContext: module.helpContext,
    cookie: module.cookie,
    readOnly: module.readOnly,
    isPrivate: module.isPrivate,
  }));

  if (isProtected && options.includeProtectedSource !== true) {
    // We surface protection; we do not defeat it. The module list stays - it is
    // in the PROJECT stream in the clear anyway - but the source does not.
    warnings.push('the project is protected; source was not extracted');
  } else {
    for (const module of modules) {
      const stream = findStream(cfb, module.streamName);
      if (stream === undefined) {
        warnings.push(`module ${module.name}: no stream named '${module.streamName}'`);
        continue;
      }
      if (module.offset > stream.length) {
        warnings.push(
          `module ${module.name}: offset ${module.offset} is past the end of a ${stream.length} byte stream`,
        );
        continue;
      }
      try {
        // Everything before MODULEOFFSET is the compiled p-code performance
        // cache. It is version specific, undocumented, and explicitly "MUST be
        // ignored on read", so it is exactly what we do with it.
        const source = decompress(stream.subarray(module.offset));
        module.source = decodeText(source, dir.codePage, warnings);
      } catch (error) {
        warnings.push(`module ${module.name}: source not decompressible (${(error as Error).message})`);
      }
    }
  }

  return {
    projectName: dir.name === '' ? properties.name : dir.name,
    codePage: dir.codePage,
    protected: isProtected,
    protection,
    modules,
    references: dir.references,
    properties,
    sysKind: dir.sysKind,
    lcid: dir.lcid,
    docString: dir.docString,
    helpFile: dir.helpFile,
    constants: dir.constants,
    versionMajor: dir.versionMajor,
    versionMinor: dir.versionMinor,
    warnings,
  };
}
