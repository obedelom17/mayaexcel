// MayaExcel – application Windows (Electron). Le site est servi localement : tout fonctionne sans Internet.
const { app, BrowserWindow, shell } = require('electron');
const http = require('http');
const fs = require('fs');
const path = require('path');

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json', '.webmanifest': 'application/manifest+json', '.png': 'image/png',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
};

function serve(root) {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      let p = decodeURIComponent(req.url.split('?')[0]);
      if (p.endsWith('/')) p += 'index.html';
      const file = path.normalize(path.join(root, p));
      if (!file.startsWith(root)) { res.writeHead(403); return res.end(); }
      fs.readFile(file, (err, data) => {
        if (err) { res.writeHead(404); return res.end('Introuvable'); }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream' });
        res.end(data);
      });
    });
    server.listen(0, '127.0.0.1', () => resolve(server.address().port));
  });
}

async function createWindow() {
  const port = await serve(path.join(__dirname, 'www'));
  const win = new BrowserWindow({
    width: 1200, height: 860, minWidth: 420, minHeight: 600,
    title: 'MayaExcel', icon: path.join(__dirname, 'icon.png'),
    autoHideMenuBar: true, backgroundColor: '#ffffff', show: false,
  });
  win.once('ready-to-show', () => { win.maximize(); win.show(); });
  win.webContents.setWindowOpenHandler(({ url }) => { shell.openExternal(url); return { action: 'deny' }; });
  win.loadURL('http://127.0.0.1:' + port + '/');
}

if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.whenReady().then(createWindow);
  app.on('window-all-closed', () => app.quit());
}
