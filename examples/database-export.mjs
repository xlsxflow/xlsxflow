// Export a database table to .xlsx without loading the table into memory. Rows are pulled from the
// query only as fast as the file is written. This uses node:sqlite (Node 22.13 and later); the same
// pattern works with any cursor or paginated query, such as pg-cursor, mysql2's stream() or a Prisma
// query with cursor pagination.
import { createWriteStream } from 'node:fs';
import { DatabaseSync } from 'node:sqlite';
import { Writable } from 'node:stream';
import { SheetWriter } from '@xlsxflow/core';

export async function exportTable(db, path) {
  const query = db.prepare('SELECT id, name, email, created_at FROM users ORDER BY id');
  async function* rows() {
    yield ['ID', 'Name', 'Email', 'Created'];
    for (const user of query.iterate()) yield [user.id, user.name, user.email, new Date(user.created_at)];
  }
  await new SheetWriter()
    .addSheet('Users', rows(), { freezePanes: { row: 1 }, autoFilter: 'A1:D1', columnWidths: [8, 20, 28, 20] })
    .write()
    .pipeTo(Writable.toWeb(createWriteStream(path)));
}

// Example: node database-export.mjs users.db users.xlsx
if (process.argv[1]?.endsWith('database-export.mjs')) {
  const [dbPath, out = 'users.xlsx'] = process.argv.slice(2);
  await exportTable(new DatabaseSync(dbPath, { readOnly: true }), out);
}
