import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { SheetWriter } from '../src/core/writer';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';
import { imageInfo } from '../src/core/image';
import { SheetReader } from '../src/core/index';

const fixture = (name: string) => new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));
const png = fixture('red-40x20.png');
const jpg = fixture('blue-30x60.jpg');
const gif = fixture('12x7.gif');

async function entries(stream: ReadableStream<Uint8Array>) {
  const zip = new ZipRandomAccessParser(createBlobReader(await new Response(stream).blob()));
  await zip.parseCentralDirectory();
  const text = async (name: string) => (zip.has(name) ? new Response(await zip.extractStream(name)).text() : undefined);
  const bytes = async (name: string) => new Uint8Array(await new Response(await zip.extractStream(name)).arrayBuffer());
  return { text, bytes, has: (name: string) => zip.has(name) };
}

describe('images', () => {
  it('reads format and size from the header', () => {
    expect(imageInfo(png)).toEqual({ ext: 'png', width: 40, height: 20 });
    expect(imageInfo(jpg)).toEqual({ ext: 'jpeg', width: 30, height: 60 }); // progressive: SOF2
    expect(imageInfo(gif)).toEqual({ ext: 'gif', width: 12, height: 7 });
    expect(() => imageInfo(new TextEncoder().encode('<svg/>'))).toThrow(/only PNG, JPEG and GIF/);
  });

  it('embeds pictures anchored to cells, storing shared data once', async () => {
    const w = new SheetWriter();
    w.addSheet('A', [['logo']], {
      images: [
        { data: png, at: 'B2' },                        // own size
        { data: jpg, at: 'D2', height: 30 },            // keeps aspect ratio: 15 x 30
        { data: png.buffer as ArrayBuffer, range: 'A5:C9', altText: 'Red & "box"' },
      ],
    });
    w.addSheet('B', [['plain']]);
    w.addSheet('C', [[{ value: 'link', hyperlink: 'https://example.com' }]], { images: [{ data: png, at: 'A1' }, { data: gif, at: 'C1' }] });
    const z = await entries(w.write());

    expect(await z.bytes('xl/media/image1.png')).toEqual(png);
    expect(z.has('xl/media/image2.jpeg')).toBe(true);
    expect(z.has('xl/media/image3.png')).toBe(true); // the ArrayBuffer is a different object
    expect(z.has('xl/media/image4.gif')).toBe(true);
    expect(z.has('xl/media/image5.png')).toBe(false);

    const d1 = (await z.text('xl/drawings/drawing1.xml'))!;
    expect(d1).toContain('<xdr:ext cx="381000" cy="190500"/>'); // 40 x 20 px
    expect(d1).toContain('<xdr:ext cx="142875" cy="285750"/>'); // 15 x 30 px
    expect(d1).toMatch(/<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0<\/xdr:col>.*?<xdr:row>4<\/xdr:row>.*?<xdr:to><xdr:col>3<\/xdr:col>.*?<xdr:row>9<\/xdr:row>/);
    expect(d1).toContain('descr="Red &amp; &quot;box&quot;"');
    expect(await z.text('xl/drawings/_rels/drawing1.xml.rels')).toContain('Target="../media/image3.png"');

    // Sheet C: its hyperlink keeps rId1, the drawing has its own id; numbering skips sheet B
    expect(await z.text('xl/worksheets/sheet3.xml')).toMatch(/<hyperlinks>.*<\/hyperlinks>\s*<drawing r:id="rIdDrawing"\/>/);
    const rels3 = (await z.text('xl/worksheets/_rels/sheet3.xml.rels'))!;
    expect(rels3).toContain('Id="rId1"');
    expect(rels3).toContain('Target="../drawings/drawing2.xml"');
    expect(await z.text('xl/worksheets/_rels/sheet2.xml.rels')).toBeUndefined();
    expect(await z.text('xl/drawings/_rels/drawing2.xml.rels')).toContain('Target="../media/image1.png"');

    const types = (await z.text('[Content_Types].xml'))!;
    for (const part of ['Extension="png"', 'Extension="jpeg"', 'Extension="gif"', '/xl/drawings/drawing1.xml', '/xl/drawings/drawing2.xml']) {
      expect(types).toContain(part);
    }
  });

  it('reads pictures back in the writer format', async () => {
    const images = [
      { data: png, at: 'B2', width: 40, height: 20 },
      { data: jpg, range: 'D2:E6', altText: 'Blue & "tall"' },
      { data: png, at: 'H10', width: 80, height: 40 },
    ];
    const w = new SheetWriter();
    w.addSheet('Pics', [['x']], { images });
    w.addSheet('None', [['y']]);
    const blob = await new Response(w.write()).blob();

    const read = await new SheetReader().parse(createBlobReader(blob), { sheetName: 'Pics' });
    const back = await read.getImages();
    expect(back).toEqual(images);
    expect(back[0].data).toBe(back[2].data); // one media file, one array
    expect(await (await new SheetReader().parse(createBlobReader(blob), { sheetName: 'None' })).getImages()).toEqual([]);
  });

  it('reads openpyxl pictures and skips charts', async () => {
    const blob = new Blob([readFileSync(new URL('../../sheetforge-pro/test/fixtures/openpyxl-chart.xlsx', import.meta.url))]);
    const back = await (await new SheetReader().parse(createBlobReader(blob))).getImages();
    expect(back).toEqual([{ data: png, at: 'A5', width: 40, height: 20, altText: 'Picture' }]);
  });
});
