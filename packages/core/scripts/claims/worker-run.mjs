// Starts worker.mjs with wrangler dev (Cloudflare's workerd) and checks its answer.
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export default async function worker(report) {
  const port = 8787 + Math.floor(Math.random() * 1000);
  const config = fileURLToPath(new URL('./wrangler.toml', import.meta.url));
  const dev = spawn(`npx -y wrangler@4.149.0 dev --config "${config}" --port ${port} --ip 127.0.0.1`, { shell: true, detached: process.platform !== 'win32', stdio: ['ignore', 'pipe', 'inherit'] });
  let log = '';
  dev.stdout.on('data', (d) => { log += d; });
  try {
    for (let i = 0; i < 120; i++) {
      await new Promise((r) => setTimeout(r, 1000));
      const res = await fetch(`http://127.0.0.1:${port}/`).catch(() => null);
      if (res) return report('Cloudflare Workers (workerd)', await res.json());
    }
    console.log(`FAIL wrangler dev did not answer\n${log}`);
    return false;
  } finally {
    if (process.platform === 'win32') spawn('taskkill', ['/pid', String(dev.pid), '/T', '/F']);
    else process.kill(-dev.pid); // the whole group: the shell, npx and wrangler
  }
}
