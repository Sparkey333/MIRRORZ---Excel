import { describe, expect, it } from 'vitest';
import { compress } from '../src/compression.js';
import { VbaProjectError, parseVbaProject } from '../src/project.js';

// ---------------------------------------------------------------------------
// There is no .xlsm in the fixture set yet, so the project under test is built
// here: a compound file containing a PROJECT stream, a compressed dir stream
// and one stream per module, assembled from the record layouts in [MS-OVBA]
// section 2.3.4.2. Building the input by hand is slower than opening a real
// file but it is the only way to assert that a specific byte in the dir stream
// comes back as a specific field, and it lets malformed projects be tested at
// all: no producer will emit one on request.
// ---------------------------------------------------------------------------

const ENDOFCHAIN = 0xffff_fffe;
const FREESECT = 0xffff_ffff;
const FATSECT = 0xffff_fffd;
const NOSTREAM = 0xffff_ffff;
const SECTOR = 512;

interface BuildNode {
  name: string;
  data?: Uint8Array;
  children?: BuildNode[];
}

interface DirEntry {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
  data?: Uint8Array;
}

function buildCfb(tree: BuildNode[]): Uint8Array {
  const entries: DirEntry[] = [
    { name: 'Root Entry', type: 5, left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, start: 0, size: 0 },
  ];

  const addLevel = (nodes: BuildNode[]): number => {
    const indices = nodes.map((node) => {
      const index = entries.length;
      entries.push({
        name: node.name,
        type: node.children === undefined ? 2 : 1,
        left: NOSTREAM,
        right: NOSTREAM,
        child: NOSTREAM,
        start: 0,
        size: 0,
        data: node.data,
      });
      return index;
    });
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i]!;
      if (node.children !== undefined && node.children.length > 0) {
        entries[indices[i]!]!.child = addLevel(node.children);
      }
      if (i + 1 < indices.length) entries[indices[i]!]!.right = indices[i + 1]!;
    }
    return indices[0]!;
  };
  entries[0]!.child = tree.length > 0 ? addLevel(tree) : NOSTREAM;

  // Small streams live in the mini stream, which is itself a normal stream
  // hanging off the root; the mini FAT chains its 64 byte sectors.
  const miniFat: number[] = [];
  const miniBytes: number[] = [];
  for (const entry of entries) {
    if (entry.data === undefined || entry.data.length === 0) continue;
    const start = miniFat.length;
    const count = Math.ceil(entry.data.length / 64);
    for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? ENDOFCHAIN : start + i + 1);
    const padded = new Uint8Array(count * 64);
    padded.set(entry.data);
    miniBytes.push(...padded);
    entry.start = start;
    entry.size = entry.data.length;
  }

  const sectors: Uint8Array[] = [];
  const fat: number[] = [];
  const allocate = (data: Uint8Array): number => {
    const count = Math.max(1, Math.ceil(data.length / SECTOR));
    const first = sectors.length;
    for (let i = 0; i < count; i++) {
      const sector = new Uint8Array(SECTOR);
      sector.set(data.subarray(i * SECTOR, (i + 1) * SECTOR));
      sectors.push(sector);
      fat.push(i === count - 1 ? ENDOFCHAIN : first + i + 1);
    }
    return first;
  };

  const miniStream = Uint8Array.from(miniBytes);
  const miniStreamStart = miniStream.length > 0 ? allocate(miniStream) : ENDOFCHAIN;
  entries[0]!.start = miniStreamStart === ENDOFCHAIN ? 0 : miniStreamStart;
  entries[0]!.size = miniStream.length;

  let firstMiniFat = ENDOFCHAIN;
  let miniFatSectors = 0;
  if (miniFat.length > 0) {
    const bytes = new Uint8Array(Math.ceil((miniFat.length * 4) / SECTOR) * SECTOR).fill(0xff);
    const view = new DataView(bytes.buffer);
    for (let i = 0; i < miniFat.length; i++) view.setUint32(i * 4, miniFat[i]!, true);
    firstMiniFat = allocate(bytes);
    miniFatSectors = bytes.length / SECTOR;
  }

  const dirBytes = new Uint8Array(Math.ceil((entries.length * 128) / SECTOR) * SECTOR);
  const dv = new DataView(dirBytes.buffer);
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]!;
    const o = i * 128;
    for (let c = 0; c < entry.name.length; c++) dv.setUint16(o + c * 2, entry.name.charCodeAt(c), true);
    dv.setUint16(o + entry.name.length * 2, 0, true);
    dv.setUint16(o + 64, entry.name.length * 2 + 2, true);
    dv.setUint8(o + 66, entry.type);
    dv.setUint8(o + 67, 1);
    dv.setUint32(o + 68, entry.left, true);
    dv.setUint32(o + 72, entry.right, true);
    dv.setUint32(o + 76, entry.child, true);
    dv.setUint32(o + 116, entry.start, true);
    dv.setUint32(o + 120, entry.size, true);
  }
  const dirStart = allocate(dirBytes);

  // The FAT has to describe the sectors it occupies, so allocating it grows it.
  const perSector = SECTOR / 4;
  let fatSectorCount = 1;
  for (;;) {
    const needed = Math.ceil((sectors.length + fatSectorCount) / perSector);
    if (needed <= fatSectorCount) break;
    fatSectorCount = needed;
  }
  const fatIds: number[] = [];
  for (let i = 0; i < fatSectorCount; i++) {
    sectors.push(new Uint8Array(SECTOR));
    fat.push(FATSECT);
    fatIds.push(sectors.length - 1);
  }
  const fatBytes = new Uint8Array(fatSectorCount * SECTOR).fill(0xff);
  const fv = new DataView(fatBytes.buffer);
  for (let i = 0; i < fat.length; i++) fv.setUint32(i * 4, fat[i]!, true);
  for (let i = 0; i < fatSectorCount; i++) {
    sectors[fatIds[i]!]!.set(fatBytes.subarray(i * SECTOR, (i + 1) * SECTOR));
  }

  const header = new Uint8Array(SECTOR);
  const hv = new DataView(header.buffer);
  header.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  hv.setUint16(24, 0x003e, true);
  hv.setUint16(26, 3, true);
  hv.setUint16(28, 0xfffe, true);
  hv.setUint16(30, 9, true);
  hv.setUint16(32, 6, true);
  hv.setUint32(40, 0, true);
  hv.setUint32(44, fatSectorCount, true);
  hv.setUint32(48, dirStart, true);
  hv.setUint32(56, 4096, true);
  hv.setUint32(60, firstMiniFat, true);
  hv.setUint32(64, miniFatSectors, true);
  hv.setUint32(68, ENDOFCHAIN, true);
  hv.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) {
    hv.setUint32(76 + i * 4, i < fatSectorCount ? fatIds[i]! : FREESECT, true);
  }

  const out = new Uint8Array(SECTOR + sectors.length * SECTOR);
  out.set(header);
  for (let i = 0; i < sectors.length; i++) out.set(sectors[i]!, SECTOR + i * SECTOR);
  return out;
}

// --- dir stream construction -------------------------------------------------

const cp1252 = (text: string): number[] => [...text].map((c) => c.charCodeAt(0) & 0xff);
const utf16 = (text: string): number[] => [...text].flatMap((c) => [c.charCodeAt(0) & 0xff, c.charCodeAt(0) >> 8]);
const u16 = (value: number): number[] => [value & 0xff, (value >> 8) & 0xff];
const u32 = (value: number): number[] => [
  value & 0xff,
  (value >>> 8) & 0xff,
  (value >>> 16) & 0xff,
  (value >>> 24) & 0xff,
];

/** id, four byte size, payload: the shape most dir records take. */
const record = (id: number, payload: number[]): number[] => [...u16(id), ...u32(payload.length), ...payload];

/** A record storing the same string as MBCS and then as UTF-16. */
const pair = (id: number, reserved: number, text: string): number[] => [
  ...u16(id),
  ...u32(text.length),
  ...cp1252(text),
  ...u16(reserved),
  ...u32(text.length * 2),
  ...utf16(text),
];

interface ModuleSpec {
  name: string;
  streamName?: string;
  source: string;
  offset?: number;
  document?: boolean;
  readOnly?: boolean;
  isPrivate?: boolean;
  docString?: string;
  extraRecords?: number[];
  omitTerminator?: boolean;
  /** Describe the module in dir but do not create its stream. */
  omitStream?: boolean;
}

interface ProjectSpec {
  name?: string;
  codePage?: number;
  modules?: ModuleSpec[];
  references?: boolean;
  projectStream?: string;
  omitProjectStream?: boolean;
  declaredModuleCount?: number;
  extraRecords?: number[];
  truncateDir?: number;
  compressDir?: boolean;
  moduleData?: (module: ModuleSpec, source: Uint8Array) => Uint8Array | undefined;
}

function dirStream(spec: ProjectSpec): Uint8Array {
  const modules = spec.modules ?? [];
  const bytes: number[] = [
    ...record(0x0001, u32(1)),
    ...record(0x0002, u32(0x0409)),
    ...record(0x0014, u32(0x0409)),
    ...record(0x0003, u16(spec.codePage ?? 1252)),
    ...record(0x0004, cp1252(spec.name ?? 'VBAProject')),
    ...pair(0x0005, 0x0040, 'Example description'),
    ...pair(0x0006, 0x003d, 'c:\\help\\example.hlp'),
    ...record(0x0007, u32(1)),
    ...record(0x0008, u32(0)),
    // PROJECTVERSION: the size field is a reserved 4 that does not cover the
    // six bytes of payload, which is the record most likely to desynchronise a
    // parser that assumes the common shape.
    ...u16(0x0009),
    ...u32(4),
    ...u32(1234),
    ...u16(7),
    ...pair(0x000c, 0x003c, 'DEBUGMODE = 1'),
  ];

  if (spec.references === true) {
    bytes.push(...pair(0x0016, 0x003e, 'stdole'));
    const libid = cp1252('*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\stdole2.tlb#OLE');
    bytes.push(
      ...record(0x000d, [...u32(libid.length), ...libid, ...u32(0), ...u16(0)]),
    );
    bytes.push(...pair(0x0016, 0x003e, 'OtherProject'));
    const absolute = cp1252('*\\C:\\books\\other.xls');
    const relative = cp1252('*\\other.xls');
    bytes.push(
      ...record(0x000e, [
        ...u32(absolute.length),
        ...absolute,
        ...u32(relative.length),
        ...relative,
        ...u32(3),
        ...u16(1),
      ]),
    );
  }

  if (spec.extraRecords !== undefined) bytes.push(...spec.extraRecords);

  bytes.push(...record(0x000f, u16(spec.declaredModuleCount ?? modules.length)));
  bytes.push(...record(0x0013, u16(0xbeef)));

  for (const module of modules) {
    bytes.push(...record(0x0019, cp1252(module.name)));
    bytes.push(...record(0x0047, utf16(module.name)));
    if (module.streamName !== undefined) bytes.push(...pair(0x001a, 0x0032, module.streamName));
    bytes.push(...pair(0x001c, 0x0048, module.docString ?? ''));
    bytes.push(...record(0x0031, u32(module.offset ?? 0)));
    bytes.push(...record(0x001e, u32(0)));
    bytes.push(...record(0x002c, u16(0xffff)));
    bytes.push(...record(module.document === true ? 0x0022 : 0x0021, []));
    if (module.readOnly === true) bytes.push(...record(0x0025, []));
    if (module.isPrivate === true) bytes.push(...record(0x0028, []));
    if (module.extraRecords !== undefined) bytes.push(...module.extraRecords);
    if (module.omitTerminator !== true) bytes.push(...record(0x002b, []));
  }

  bytes.push(...record(0x0010, []));

  const raw = Uint8Array.from(spec.truncateDir === undefined ? bytes : bytes.slice(0, spec.truncateDir));
  return spec.compressDir === false ? raw : compress(raw);
}

const DEFAULT_PROJECT_STREAM = [
  'ID="{917DED54-440B-4FD1-A5C1-74ACF261E600}"',
  'Document=ThisWorkbook/&H00000000',
  'Document=Sheet1/&H00000000',
  'Module=Module1',
  'Class=Class1',
  'BaseClass=UserForm1',
  'HelpFile="c:\\help\\example.hlp"',
  'Name="VBAProject"',
  'HelpContextID="1"',
  'Description="Example description"',
  'VersionCompatible32="393222000"',
  'CMG="0705D8E3D8EDDBF1DBF1DBF1DBF1"',
  'DPB="0E0CD1ECDFF4E7F5E7F5E7"',
  'GC="1517CAF1D6F9D7F9D706"',
  '',
  '[Host Extender Info]',
  '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
  '',
  '[Workspace]',
  'Sheet1=69, 69, 724, 317, C',
].join('\r\n');

function buildProject(spec: ProjectSpec = {}): Uint8Array {
  const modules = spec.modules ?? [];
  const vbaChildren: BuildNode[] = [{ name: 'dir', data: dirStream(spec) }];
  for (const module of modules) {
    const offset = module.offset ?? 0;
    // Everything before MODULEOFFSET is the p-code cache, which we fill with
    // bytes that would decompress to nonsense if anyone tried.
    const cache = new Uint8Array(offset).fill(0xcc);
    const source = compress(new TextEncoder().encode(module.source));
    if (module.omitStream === true) continue;
    const stream = spec.moduleData?.(module, source) ?? Uint8Array.from([...cache, ...source]);
    vbaChildren.push({ name: module.streamName ?? module.name, data: stream });
  }
  const root: BuildNode[] = [];
  if (spec.omitProjectStream !== true) {
    root.push({
      name: 'PROJECT',
      data: Uint8Array.from(cp1252(spec.projectStream ?? DEFAULT_PROJECT_STREAM)),
    });
  }
  root.push({ name: 'VBA', children: vbaChildren });
  return buildCfb(root);
}

const HELLO: ModuleSpec = {
  name: 'Module1',
  streamName: 'Module1',
  source: 'Option Explicit\r\n\r\nSub Hello()\r\n  MsgBox "hello"\r\nEnd Sub\r\n',
  offset: 40,
};

// ---------------------------------------------------------------------------

describe('parseVbaProject: a well-formed project', () => {
  const project = parseVbaProject(
    buildProject({ modules: [HELLO], references: true }),
  );

  it('reports no warnings', () => {
    expect(project.warnings).toEqual([]);
  });

  it('reads the project name from the dir stream', () => {
    expect(project.projectName).toBe('VBAProject');
  });

  it('reads the code page', () => {
    expect(project.codePage).toBe(1252);
  });

  it('reads the platform, locale and version', () => {
    expect(project.sysKind).toBe(1);
    expect(project.lcid).toBe(0x0409);
    expect(project.versionMajor).toBe(1234);
    expect(project.versionMinor).toBe(7);
  });

  it('reads the project description, help file and constants', () => {
    expect(project.docString).toBe('Example description');
    expect(project.helpFile).toBe('c:\\help\\example.hlp');
    expect(project.constants).toBe('DEBUGMODE = 1');
  });

  it('finds one module', () => {
    expect(project.modules.map((m) => m.name)).toEqual(['Module1']);
  });

  it('reads the module stream name and offset', () => {
    expect(project.modules[0]!.streamName).toBe('Module1');
    expect(project.modules[0]!.offset).toBe(40);
  });

  it('decompresses the module source from MODULEOFFSET onwards', () => {
    expect(project.modules[0]!.source).toBe(HELLO.source);
  });

  it('classifies a procedural module', () => {
    expect(project.modules[0]!.type).toBe('procedural');
  });

  it('is not protected', () => {
    expect(project.protected).toBe(false);
    expect(project.protection).toEqual({
      userProtected: false,
      hostProtected: false,
      vbeProtected: false,
      passwordSet: false,
    });
  });

  it('reads a registered reference with its name', () => {
    const registered = project.references.filter((r) => r.kind === 'registered');
    expect(registered).toHaveLength(1);
    expect(registered[0]!.name).toBe('stdole');
    expect(registered[0]!.libid).toContain('{00020430-0000-0000-C000-000000000046}');
  });

  it('reads a project reference with both paths and its version', () => {
    const referenced = project.references.find((r) => r.kind === 'project')!;
    expect(referenced.name).toBe('OtherProject');
    expect(referenced.libid).toBe('*\\C:\\books\\other.xls');
    expect(referenced.libidRelative).toBe('*\\other.xls');
    expect(referenced.majorVersion).toBe(3);
    expect(referenced.minorVersion).toBe(1);
  });
});

describe('parseVbaProject: several modules', () => {
  const modules: ModuleSpec[] = [
    { name: 'Module1', streamName: 'Module1', source: 'Sub A()\r\nEnd Sub\r\n', offset: 12 },
    {
      name: 'ThisWorkbook',
      streamName: 'ThisWorkbook',
      source: 'Private Sub Workbook_Open()\r\nEnd Sub\r\n',
      offset: 0,
      document: true,
    },
    {
      name: 'Sheet1',
      streamName: 'Sheet1',
      source: '',
      offset: 100,
      document: true,
      readOnly: true,
      isPrivate: true,
      docString: 'the first sheet',
    },
  ];
  const project = parseVbaProject(buildProject({ modules }));

  it('keeps the modules in dir order', () => {
    expect(project.modules.map((m) => m.name)).toEqual(['Module1', 'ThisWorkbook', 'Sheet1']);
  });

  it('reads every source', () => {
    expect(project.modules.map((m) => m.source)).toEqual([
      'Sub A()\r\nEnd Sub\r\n',
      'Private Sub Workbook_Open()\r\nEnd Sub\r\n',
      '',
    ]);
  });

  it('distinguishes document modules from procedural ones', () => {
    expect(project.modules.map((m) => m.type)).toEqual(['procedural', 'document', 'document']);
  });

  it('reads the read-only and private flags', () => {
    expect(project.modules[2]!.readOnly).toBe(true);
    expect(project.modules[2]!.isPrivate).toBe(true);
    expect(project.modules[0]!.readOnly).toBe(false);
    expect(project.modules[0]!.isPrivate).toBe(false);
  });

  it('reads the module description', () => {
    expect(project.modules[2]!.docString).toBe('the first sheet');
  });

  it('reads the module cookie and help context', () => {
    expect(project.modules[0]!.cookie).toBe(0xffff);
    expect(project.modules[0]!.helpContext).toBe(0);
  });

  it('handles a zero MODULEOFFSET, where there is no p-code cache at all', () => {
    expect(project.modules[1]!.offset).toBe(0);
    expect(project.modules[1]!.source).toContain('Workbook_Open');
  });
});

describe('parseVbaProject: the PROJECT stream', () => {
  const project = parseVbaProject(buildProject({ modules: [HELLO] }));

  it('splits the module kinds out of the property list', () => {
    expect(project.properties.documents).toEqual(['ThisWorkbook', 'Sheet1']);
    expect(project.properties.modules).toEqual(['Module1']);
    expect(project.properties.classes).toEqual(['Class1']);
    expect(project.properties.baseClasses).toEqual(['UserForm1']);
  });

  it('strips the quotes from quoted values', () => {
    expect(project.properties.name).toBe('VBAProject');
    expect(project.properties.description).toBe('Example description');
    expect(project.properties.helpFile).toBe('c:\\help\\example.hlp');
    expect(project.properties.helpContextId).toBe('1');
    expect(project.properties.versionCompatible32).toBe('393222000');
  });

  it('keeps the project identifier', () => {
    expect(project.properties.id).toBe('{917DED54-440B-4FD1-A5C1-74ACF261E600}');
  });

  it('stops at the first section header', () => {
    const keys = project.properties.entries.map(([key]) => key);
    expect(keys).toContain('GC');
    expect(keys.some((key) => key.startsWith('&H'))).toBe(false);
    expect(keys).not.toContain('Sheet1');
  });

  it('keeps the raw text for anything we did not model', () => {
    expect(project.properties.text).toContain('[Workspace]');
  });

  it('survives a PROJECT stream with blank lines and no properties', () => {
    const project = parseVbaProject(buildProject({ modules: [HELLO], projectStream: '\r\n\r\n' }));
    expect(project.properties.entries).toEqual([]);
    expect(project.projectName).toBe('VBAProject');
  });

  it('ignores a line with no equals sign', () => {
    const project = parseVbaProject(
      buildProject({ modules: [HELLO], projectStream: 'nonsense\r\nName="Kept"\r\n' }),
    );
    expect(project.properties.name).toBe('Kept');
  });

  it('warns but carries on when the PROJECT stream is missing', () => {
    const project = parseVbaProject(buildProject({ modules: [HELLO], omitProjectStream: true }));
    expect(project.warnings.join(' ')).toContain("no 'PROJECT' stream");
    expect(project.modules[0]!.source).toBe(HELLO.source);
  });
});

describe('parseVbaProject: protection', () => {
  /**
   * Fixture-only. This writes the obfuscation of section 2.4.3 so that a
   * protected project can be built to test against; it is not in the library
   * and it is never used to write a password. The value it produces here is a
   * protection-state word, which is exactly what we read back.
   */
  const obfuscate = (data: number[], seed: number): string => {
    const projKey = 0x42;
    const versionEnc = seed ^ 2;
    const projKeyEnc = seed ^ projKey;
    let unencrypted = projKey;
    let enc1 = projKeyEnc;
    let enc2 = versionEnc;
    const out = [seed, versionEnc, projKeyEnc];
    const push = (plain: number): void => {
      const byteEnc = (plain ^ (enc2 + unencrypted)) & 0xff;
      out.push(byteEnc);
      enc2 = enc1;
      enc1 = byteEnc;
      unencrypted = plain;
    };
    for (let i = 0; i < (seed & 6) / 2; i++) push(0x07);
    for (const byte of u32(data.length)) push(byte);
    for (const byte of data) push(byte);
    return out.map((b) => b.toString(16).padStart(2, '0').toUpperCase()).join('');
  };

  const withProtection = (cmg: number[], dpb: number[]): Uint8Array =>
    buildProject({
      modules: [HELLO],
      projectStream: [
        'ID="{917DED54-440B-4FD1-A5C1-74ACF261E600}"',
        'Name="VBAProject"',
        `CMG="${obfuscate(cmg, 0x07)}"`,
        `DPB="${obfuscate(dpb, 0x0e)}"`,
        `GC="${obfuscate([0xff], 0x15)}"`,
      ].join('\r\n'),
    });

  it('decodes the unprotected CMG and DPB from the specification example', () => {
    // The published example in section 3.1.6 is an unprotected project, so
    // these exact hex strings must come back as "not protected".
    const project = parseVbaProject(buildProject({ modules: [HELLO] }));
    expect(project.protected).toBe(false);
    expect(project.protection.passwordSet).toBe(false);
  });

  it('reports a user-locked project', () => {
    const project = parseVbaProject(withProtection([0x01, 0, 0, 0], [0x00]));
    expect(project.protection.userProtected).toBe(true);
    expect(project.protected).toBe(true);
  });

  it('reports host protection', () => {
    const project = parseVbaProject(withProtection([0x02, 0, 0, 0], [0x00]));
    expect(project.protection.hostProtected).toBe(true);
    expect(project.protection.userProtected).toBe(false);
  });

  it('reports editor protection', () => {
    const project = parseVbaProject(withProtection([0x04, 0, 0, 0], [0x00]));
    expect(project.protection.vbeProtected).toBe(true);
  });

  it('reports a password from the length of the DPB payload alone', () => {
    const project = parseVbaProject(withProtection([0x00, 0, 0, 0], [1, 2, 3, 4, 5, 6, 7, 8]));
    expect(project.protection.passwordSet).toBe(true);
    expect(project.protected).toBe(true);
  });

  it('withholds source from a protected project and says so', () => {
    const project = parseVbaProject(withProtection([0x01, 0, 0, 0], [0x00]));
    expect(project.modules[0]!.source).toBe('');
    expect(project.warnings.join(' ')).toContain('protected');
  });

  it('still lists the modules of a protected project', () => {
    const project = parseVbaProject(withProtection([0x01, 0, 0, 0], [0x00]));
    expect(project.modules.map((m) => m.name)).toEqual(['Module1']);
    expect(project.modules[0]!.offset).toBe(40);
  });

  it('extracts source from a protected project only when explicitly asked', () => {
    const project = parseVbaProject(withProtection([0x01, 0, 0, 0], [0x00]), {
      includeProtectedSource: true,
    });
    expect(project.modules[0]!.source).toBe(HELLO.source);
    expect(project.protected).toBe(true);
  });

  it('treats an unparsable protection blob as unprotected rather than failing', () => {
    const project = parseVbaProject(
      buildProject({ modules: [HELLO], projectStream: 'CMG="not hex"\r\nDPB="0E0C"\r\n' }),
    );
    expect(project.protected).toBe(false);
  });

  it('treats a project with no CMG or DPB at all as unprotected', () => {
    const project = parseVbaProject(
      buildProject({ modules: [HELLO], projectStream: 'Name="Bare"\r\n' }),
    );
    expect(project.protected).toBe(false);
  });
});

describe('parseVbaProject: code pages', () => {
  it('decodes windows-1252 high bytes in module source', () => {
    const source = 'Rem \u2018smart quotes\u2019 and an em dash \u2014 here\r\n';
    const project = parseVbaProject(
      buildProject({
        codePage: 1252,
        modules: [{ name: 'M', streamName: 'M', source: '', offset: 0 }],
        moduleData: () => {
          const bytes = [...source].map((c) => {
            if (c === '\u2018') return 0x91;
            if (c === '\u2019') return 0x92;
            if (c === '\u2014') return 0x97;
            return c.charCodeAt(0);
          });
          return compress(Uint8Array.from(bytes));
        },
      }),
    );
    expect(project.modules[0]!.source).toBe(source);
  });

  it('decodes utf-8 source when the code page is 65001', () => {
    const source = 'Rem \u00e9\u00e8 \u4f60\u597d\r\n';
    const project = parseVbaProject(
      buildProject({
        codePage: 65001,
        modules: [{ name: 'M', streamName: 'M', source, offset: 0 }],
      }),
    );
    expect(project.codePage).toBe(65001);
    expect(project.modules[0]!.source).toBe(source);
  });

  it('falls back to windows-1252 and warns for a code page it does not implement', () => {
    const project = parseVbaProject(
      buildProject({
        codePage: 1361,
        modules: [{ name: 'M', streamName: 'M', source: 'Sub A()\r\nEnd Sub\r\n', offset: 0 }],
      }),
    );
    expect(project.modules[0]!.source).toBe('Sub A()\r\nEnd Sub\r\n');
    expect(project.warnings.join(' ')).toMatch(/code page 1361/);
  });

  it('does not crash on a code page of zero', () => {
    const project = parseVbaProject(
      buildProject({ codePage: 0, modules: [{ name: 'M', streamName: 'M', source: 'x', offset: 0 }] }),
    );
    expect(project.modules[0]!.source).toBe('x');
  });
});

describe('parseVbaProject: malformed input', () => {
  it('rejects a buffer that is not a compound file', () => {
    expect(() => parseVbaProject(new Uint8Array(1024))).toThrow();
  });

  it('rejects a compound file with no dir stream', () => {
    const cfb = buildCfb([{ name: 'PROJECT', data: Uint8Array.from(cp1252('Name="X"')) }]);
    expect(() => parseVbaProject(cfb)).toThrow(VbaProjectError);
  });

  it('rejects a dir stream that is not compressed', () => {
    expect(() => parseVbaProject(buildProject({ modules: [HELLO], compressDir: false }))).toThrow(
      /not decompressible/,
    );
  });

  it('skips an unknown project-level record and carries on', () => {
    const project = parseVbaProject(
      buildProject({ modules: [HELLO], extraRecords: record(0x0099, [1, 2, 3, 4, 5]) }),
    );
    expect(project.warnings.join(' ')).toContain('0x0099');
    expect(project.modules[0]!.source).toBe(HELLO.source);
  });

  it('skips an unknown module-level record and carries on', () => {
    const project = parseVbaProject(
      buildProject({ modules: [{ ...HELLO, extraRecords: record(0x00fe, [9, 9]) }] }),
    );
    expect(project.warnings.join(' ')).toContain('0x00fe');
    expect(project.modules[0]!.name).toBe('Module1');
    expect(project.modules[0]!.source).toBe(HELLO.source);
  });

  it('stops early and warns when the dir stream is truncated mid-record', () => {
    const project = parseVbaProject(buildProject({ modules: [HELLO], truncateDir: 25 }));
    expect(project.warnings.join(' ')).toContain('stopped early');
    expect(project.modules).toEqual([]);
  });

  it('warns when fewer modules are present than PROJECTMODULES declared', () => {
    const project = parseVbaProject(buildProject({ modules: [HELLO], declaredModuleCount: 4 }));
    expect(project.warnings.join(' ')).toContain('declared 4 modules');
    expect(project.modules).toHaveLength(1);
  });

  it('warns when more modules are present than PROJECTMODULES declared', () => {
    const project = parseVbaProject(
      buildProject({
        modules: [HELLO, { name: 'Module2', streamName: 'Module2', source: 'Sub B()\r\nEnd Sub\r\n' }],
        declaredModuleCount: 1,
      }),
    );
    expect(project.warnings.join(' ')).toContain('declared count');
  });

  it('warns when a module stream is missing altogether', () => {
    const project = parseVbaProject(buildProject({ modules: [{ ...HELLO, omitStream: true }] }));
    expect(project.warnings.join(' ')).toContain("no stream named 'Module1'");
    expect(project.modules[0]!.source).toBe('');
  });

  it('warns when a module stream is present but empty', () => {
    const project = parseVbaProject(
      buildProject({ modules: [HELLO], moduleData: () => new Uint8Array(0) }),
    );
    expect(project.warnings.join(' ')).toContain('past the end');
    expect(project.modules[0]!.source).toBe('');
  });

  it('warns when MODULEOFFSET points past the end of its stream', () => {
    const project = parseVbaProject(
      buildProject({ modules: [{ ...HELLO, offset: 5000 }], moduleData: (_m, source) => source }),
    );
    expect(project.warnings.join(' ')).toContain('past the end');
    expect(project.modules[0]!.source).toBe('');
  });

  it('warns when the source at MODULEOFFSET is not a compressed container', () => {
    const project = parseVbaProject(
      buildProject({
        modules: [{ ...HELLO, offset: 0 }],
        moduleData: () => Uint8Array.from([0x99, 0x00, 0x00, 0x00]),
      }),
    );
    expect(project.warnings.join(' ')).toContain('not decompressible');
    expect(project.modules[0]!.source).toBe('');
  });

  it('keeps going when one module of three is broken', () => {
    const good: ModuleSpec = { name: 'Good', streamName: 'Good', source: 'Sub G()\r\nEnd Sub\r\n' };
    const bad: ModuleSpec = { name: 'Bad', streamName: 'Bad', source: 'Sub B()\r\nEnd Sub\r\n' };
    const project = parseVbaProject(
      buildProject({
        modules: [good, bad, { ...good, name: 'Good2', streamName: 'Good2' }],
        moduleData: (module, source) =>
          module.name === 'Bad' ? Uint8Array.from([0x02, 0x03]) : source,
      }),
    );
    expect(project.modules.map((m) => m.source !== '')).toEqual([true, false, true]);
  });

  it('falls back to the module name when MODULESTREAMNAME is absent', () => {
    const project = parseVbaProject(
      buildProject({ modules: [{ name: 'Module1', source: 'Sub A()\r\nEnd Sub\r\n' }] }),
    );
    expect(project.modules[0]!.streamName).toBe('Module1');
    expect(project.modules[0]!.source).toBe('Sub A()\r\nEnd Sub\r\n');
  });

  it('stops the module walk cleanly when a module record has no terminator', () => {
    const project = parseVbaProject(
      buildProject({ modules: [{ ...HELLO, omitTerminator: true }] }),
    );
    // The dir terminator is consumed as part of the unterminated module, so the
    // module list may be empty; what matters is that nothing threw.
    expect(project.projectName).toBe('VBAProject');
  });

  it('accepts a project with no modules at all', () => {
    const project = parseVbaProject(buildProject({ modules: [] }));
    expect(project.modules).toEqual([]);
    expect(project.warnings).toEqual([]);
  });

  it('reads a module whose source is a single empty compressed container', () => {
    const project = parseVbaProject(
      buildProject({ modules: [{ name: 'M', streamName: 'M', source: '', offset: 0 }] }),
    );
    expect(project.modules[0]!.source).toBe('');
  });

  it('reads source spanning several compression chunks', () => {
    const source = 'Sub Big()\r\n' + '  Debug.Print "line"\r\n'.repeat(600) + 'End Sub\r\n';
    const project = parseVbaProject(
      buildProject({ modules: [{ name: 'Big', streamName: 'Big', source, offset: 17 }] }),
    );
    expect(project.modules[0]!.source).toBe(source);
    expect(project.modules[0]!.source.length).toBeGreaterThan(4096);
  });
});
