const http = require('http');
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');

const PORT = 8088;
const ROOT = path.join(__dirname);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml'
};

const server = http.createServer((req, res) => {
  let reqPath = decodeURI(req.url.split('?')[0]);
  if (reqPath === '/') reqPath = '/index.html';
  const filePath = path.join(ROOT, reqPath);

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found: ' + reqPath);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, async () => {
  console.log('Server running on http://localhost:' + PORT);

  const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
  const tasks = [
    { url: `http://localhost:${PORT}/task1.html`, out: path.join(ROOT, 'screenshots', 'task1_1280px.png'), width: 1280, height: 950 },
    { url: `http://localhost:${PORT}/task1.html`, out: path.join(ROOT, 'screenshots', 'task1_375px.png'), width: 375, height: 1150 },
    { url: `http://localhost:${PORT}/task2.html`, out: path.join(ROOT, 'screenshots', 'task2_1280px.png'), width: 1280, height: 1200 },
    { url: `http://localhost:${PORT}/task3.html`, out: path.join(ROOT, 'screenshots', 'task3_1280px.png'), width: 1280, height: 1200 },
    { url: `http://localhost:${PORT}/index.html`, out: path.join(ROOT, 'screenshots', 'index_1280px.png'), width: 1280, height: 1050 }
  ];

  for (const t of tasks) {
    await new Promise((resolve) => {
      console.log(`Capturing ${t.url} (${t.width}x${t.height}) -> ${t.out}`);
      const args = [
        '--headless=new',
        `--screenshot=${t.out}`,
        `--window-size=${t.width},${t.height}`,
        '--hide-scrollbars',
        '--disable-gpu',
        t.url
      ];
      execFile(edgePath, args, (err, stdout, stderr) => {
        if (err) console.error('Error for ' + t.out, err);
        else console.log('Saved: ' + t.out);
        resolve();
      });
    });
  }

  console.log('All screenshots completed successfully!');
  server.close(() => process.exit(0));
});
