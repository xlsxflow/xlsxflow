// Send a generated .xlsx from Express, Fastify (reply.raw) or plain node:http.
// The file streams to the client while it is written, and stops early if the client goes away.
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { SheetWriter } from '@xlsxflow/core';

async function* rows() {
  yield ['Day', 'Visitors'];
  for (let day = 1; day <= 365; day++) yield [new Date(Date.UTC(2026, 0, day)), 1000 + ((day * 37) % 500)];
}

export function sendReport(req, res) {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="visitors.xlsx"');
  // pipeline, unlike pipe, cancels the writer (and ends rows()) when the client disconnects
  pipeline(Readable.fromWeb(new SheetWriter().addSheet('Visitors', rows()).write()), res).catch(() => {});
}

// Express: app.get('/report', sendReport)
// node:http: http.createServer(sendReport).listen(3000)
