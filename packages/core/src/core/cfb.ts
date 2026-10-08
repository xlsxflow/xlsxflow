// Reader for Compound File Binary files [MS-CFB]: the container of .xls workbooks and of
// password-protected .xlsx files. Streams are read from an in-memory copy of the file.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const END_OF_CHAIN = 0xfffffffe;
const MAX_REGULAR_SECTOR = 0xfffffffa;

export interface CfbEntry {
  path: string; // "Workbook", or "Storage/Stream" inside a storage
  type: 'storage' | 'stream';
  size: number;
}

export function isCfb(bytes: Uint8Array): boolean {
  return bytes.length >= 8 && SIGNATURE.every((b, i) => bytes[i] === b);
}

const corrupt = (why: string) => new Error(`Corrupt compound file: ${why}`);

interface DirEntry { name: string; type: number; left: number; right: number; child: number; start: number; size: number }

export class CfbReader {
  private readonly view: DataView;
  private readonly sectorSize: number;
  private readonly miniSectorSize: number;
  private readonly miniCutoff: number;
  private readonly fat: Uint32Array;
  private miniFat?: Uint32Array;
  private miniStream?: Uint8Array;
  private readonly dir: DirEntry[];
  private readonly byPath = new Map<string, DirEntry & { path: string }>();

  constructor(private readonly bytes: Uint8Array) {
    if (!isCfb(bytes) || bytes.length < 512) throw new Error('Not a compound file (no D0CF11E0 signature)');
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const u16 = (o: number) => this.view.getUint16(o, true);
    const u32 = (o: number) => this.view.getUint32(o, true);

    const sectorShift = u16(0x1e);
    if (sectorShift !== 9 && sectorShift !== 12) throw corrupt(`sector shift ${sectorShift}`);
    this.sectorSize = 1 << sectorShift;
    const miniShift = u16(0x20);
    if (miniShift !== 6) throw corrupt(`mini sector shift ${miniShift}`);
    this.miniSectorSize = 1 << miniShift;
    this.miniCutoff = u32(0x38);

    // The FAT's own sectors are listed in the header (109 slots) and then in a chain of DIFAT sectors
    const numFatSectors = u32(0x2c);
    const sectorCount = Math.ceil(bytes.length / this.sectorSize) - 1;
    if (numFatSectors > sectorCount) throw corrupt(`${numFatSectors} FAT sectors in a ${bytes.length}-byte file`);
    const fatSectors: number[] = [];
    for (let i = 0; i < 109 && fatSectors.length < numFatSectors; i++) fatSectors.push(u32(0x4c + i * 4));
    const perDifat = this.sectorSize / 4 - 1;
    for (let s = u32(0x44), hops = 0; fatSectors.length < numFatSectors; hops++) {
      if (s > MAX_REGULAR_SECTOR || hops > sectorCount) throw corrupt('DIFAT chain ends early');
      const at = this.sectorOffset(s);
      for (let i = 0; i < perDifat && fatSectors.length < numFatSectors; i++) fatSectors.push(u32(at + i * 4));
      s = u32(at + perDifat * 4);
    }
    const perSector = this.sectorSize / 4;
    this.fat = new Uint32Array(numFatSectors * perSector);
    fatSectors.forEach((s, i) => {
      const at = this.sectorOffset(s);
      for (let j = 0; j < perSector; j++) this.fat[i * perSector + j] = u32(at + j * 4);
    });

    // Directory: 128-byte entries in the chain starting at the header's first directory sector
    const dirBytes = this.readChain(u32(0x30), Infinity);
    this.dir = [];
    for (let at = 0; at + 128 <= dirBytes.length; at += 128) {
      const d = new DataView(dirBytes.buffer, dirBytes.byteOffset + at, 128);
      const nameLen = Math.min(d.getUint16(64, true), 64);
      let name = '';
      for (let i = 0; i + 2 < nameLen; i += 2) name += String.fromCharCode(d.getUint16(i, true));
      // Version 3 files may leave garbage in the size's high half
      const size = sectorShift === 9 ? d.getUint32(120, true) : d.getUint32(120, true) + d.getUint32(124, true) * 2 ** 32;
      this.dir.push({ name, type: d.getUint8(66), left: d.getUint32(68, true), right: d.getUint32(72, true),
        child: d.getUint32(76, true), start: d.getUint32(116, true), size });
    }
    if (this.dir[0]?.type !== 5) throw corrupt('no root entry');
    this.walk(this.dir[0].child, '');
  }

  // Start of a sector, checked to hold `need` bytes (the last sector of a file may be cut short)
  private sectorOffset(sector: number, need = this.sectorSize): number {
    const at = (sector + 1) * this.sectorSize;
    if (sector > MAX_REGULAR_SECTOR || at + need > this.bytes.length) {
      throw corrupt(`sector ${sector} is outside the file`);
    }
    return at;
  }

  // Sibling entries form a red-black tree; children of a storage hang off its `child`
  private walk(root: number, prefix: string) {
    const stack = [root];
    const seen = new Set<number>();
    while (stack.length) {
      const id = stack.pop()!;
      if (id > MAX_REGULAR_SECTOR) continue; // NOSTREAM
      if (seen.has(id) || !this.dir[id]) throw corrupt('directory tree loops or points outside the directory');
      seen.add(id);
      const e = this.dir[id];
      stack.push(e.left, e.right);
      if (e.type !== 1 && e.type !== 2) continue;
      const path = prefix + e.name;
      this.byPath.set(path.toLowerCase(), { ...e, path });
      if (e.type === 1) this.walk(e.child, path + '/');
    }
  }

  private readChain(start: number, size: number, mini = false): Uint8Array {
    const table = mini ? this.miniFat! : this.fat;
    const unit = mini ? this.miniSectorSize : this.sectorSize;
    const source = mini ? this.miniStream! : this.bytes;
    const sectors: number[] = [];
    for (let s = start; s !== END_OF_CHAIN && sectors.length * unit < size; s = table[s]) {
      if (s >= table.length || sectors.length > table.length) throw corrupt(`${mini ? 'mini ' : ''}sector chain is broken or loops`);
      sectors.push(s);
    }
    const total = Math.min(size, sectors.length * unit);
    if (size !== Infinity && total < size) throw corrupt(`stream is shorter (${total} bytes) than its stated ${size}`);
    const out = new Uint8Array(total);
    sectors.forEach((s, i) => {
      const len = Math.min(unit, total - i * unit);
      const at = mini ? s * unit : this.sectorOffset(s, len);
      if (at + len > source.length) throw corrupt(`mini sector ${s} is outside the mini stream`);
      out.set(source.subarray(at, at + len), i * unit);
    });
    return out;
  }

  entries(): CfbEntry[] {
    return [...this.byPath.values()].map(e => ({ path: e.path, type: e.type === 1 ? 'storage' : 'stream', size: e.size }));
  }

  has(path: string): boolean {
    return this.byPath.get(path.toLowerCase())?.type === 2;
  }

  // A stream's bytes, by path (case-insensitive, as in the format); undefined if there is none
  read(path: string): Uint8Array | undefined {
    const e = this.byPath.get(path.toLowerCase());
    if (!e || e.type !== 2) return undefined;
    if (e.size >= this.miniCutoff) {
      if (e.size > this.bytes.length) throw corrupt(`stream ${e.path} is larger than the file`);
      return this.readChain(e.start, e.size);
    }
    if (!this.miniStream) {
      const root = this.dir[0];
      this.miniStream = this.readChain(root.start, root.size);
      const view = new DataView(this.bytes.buffer, this.bytes.byteOffset, this.bytes.byteLength);
      const miniFatBytes = this.readChain(view.getUint32(0x3c, true), view.getUint32(0x40, true) * this.sectorSize);
      this.miniFat = new Uint32Array(miniFatBytes.length / 4);
      const mv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
      for (let i = 0; i < this.miniFat.length; i++) this.miniFat[i] = mv.getUint32(i * 4, true);
    }
    return this.readChain(e.start, e.size, true);
  }
}
