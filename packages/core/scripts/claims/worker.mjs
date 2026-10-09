// Runs claims.mjs inside workerd (Cloudflare Workers), without nodejs_compat. GET / answers with the results.
import * as lib from '../../dist/index.mjs';
import { run } from './claims.mjs';
import xls from '../../test/fixtures/excel-made.xls';
import ods from '../../test/fixtures/excel-made.ods';
import xlsx from '../../test/fixtures/excel-made.xlsx';
import png from '../../test/fixtures/red-40x20.png';
import jpg from '../../test/fixtures/blue-30x60.jpg';
import gif from '../../test/fixtures/12x7.gif';
import encrypted from '../../test/fixtures/encrypted.xlsx';

export default {
  async fetch() {
    const fx = Object.fromEntries(Object.entries({ xls, ods, xlsx, png, jpg, gif, encrypted }).map(([k, v]) => [k, new Uint8Array(v)]));
    return Response.json(await run(lib, fx));
  },
};
