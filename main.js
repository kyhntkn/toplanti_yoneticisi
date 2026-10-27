const { app, BrowserWindow, Tray, Menu, ipcMain, shell, session, Notification, nativeImage } = require('electron');
const path = require('path'), fs = require('fs'), os = require('os'), { spawn } = require('child_process');

app.setAppUserModelId('com.toplanti.yoneticisi');
if (!app.requestSingleInstanceLock()) app.quit();

let win, tray, quitting = false, hinted = false;
const icon = () => nativeImage.createFromPath(path.join(__dirname, 'assets', 'icon.png'));

function createWindow() {
  win = new BrowserWindow({
    width: 1280, height: 820, minWidth: 900, minHeight: 600,
    title: 'Toplantı Yöneticisi', icon: icon(), show: false, backgroundColor: '#f4f6fa',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true, nodeIntegration: false, sandbox: true,
      backgroundThrottling: false   // pencere gizliyken bile bildirim zamanlayıcısı çalışsın
    }
  });
  win.loadFile(path.join(__dirname, 'app', 'index.html'));
  win.once('ready-to-show', () => { if (!process.argv.includes('--hidden')) win.show(); });
  const external = (u) => { if (/^(https?|mailto):/i.test(u)) shell.openExternal(u); };
  win.webContents.setWindowOpenHandler(({ url }) => { external(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('file:')) { e.preventDefault(); external(url); } });
  // Pencereyi kapatınca sistem tepsisine gizle: bildirimler çalışmaya devam eder
  win.on('close', (e) => {
    if (quitting) return;
    e.preventDefault(); win.hide();
    if (!hinted) {
      hinted = true;
      new Notification({ title: 'Toplantı Yöneticisi', body: 'Uygulama sistem tepsisinde çalışmaya devam ediyor; hatırlatmalar gelmeye devam eder.' }).show();
    }
  });
}

function createTray() {
  tray = new Tray(icon().resize({ width: 16, height: 16 }));
  tray.setToolTip('Toplantı Yöneticisi');
  const show = () => { win.show(); win.focus(); };
  tray.setContextMenu(Menu.buildFromTemplate([
    { label: 'Aç', click: show },
    { label: 'Windows ile başlat', type: 'checkbox', checked: app.getLoginItemSettings().openAtLogin,
      click: (i) => app.setLoginItemSettings({ openAtLogin: i.checked, args: ['--hidden'] }) },
    { type: 'separator' },
    { label: 'Çıkış', click: () => { quitting = true; app.quit(); } }
  ]));
  tray.on('click', show);
}

app.on('second-instance', () => { if (win) { win.show(); win.focus(); } });
app.on('before-quit', () => { quitting = true; });
app.on('window-all-closed', () => {});   // tepside çalışmaya devam et

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  session.defaultSession.setPermissionRequestHandler((wc, p, cb) => cb(p === 'notifications'));
  session.defaultSession.setPermissionCheckHandler((wc, p) => p === 'notifications');
  createWindow(); createTray();
});

/* ---- Outlook ile otomatik e-posta (klasik Outlook, COM) ---- */
const PS = `param([string]$job)
$ErrorActionPreference='Stop'
$j=Get-Content -Raw -Encoding UTF8 -LiteralPath $job | ConvertFrom-Json
try{$ol=[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')}catch{$ol=New-Object -ComObject Outlook.Application}
$m=$ol.CreateItem(0)
$m.To=[string]$j.to
if($j.cc){$m.CC=[string]$j.cc}
$m.Subject=[string]$j.subject
$m.Body=[string]$j.body
foreach($f in @($j.files)){if($f){[void]$m.Attachments.Add([string]$f)}}
if($j.send){$m.Send()}else{$m.Display()}
`;

ipcMain.handle('mail', (e, o) => new Promise((res) => {
  if (process.platform !== 'win32') return res({ ok: false, error: 'Yalnızca Windows ve klasik Outlook ile çalışır.' });
  let dir;
  try {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tm-'));
    const files = (o.attachments || []).map((a) => {
      const p = path.join(dir, String(a.name).replace(/[\\/:*?"<>|]/g, '_'));
      fs.writeFileSync(p, Buffer.from(a.data)); return p;
    });
    fs.writeFileSync(path.join(dir, 'job.json'),
      JSON.stringify({ to: o.to || '', cc: o.cc || '', subject: o.subject || '', body: o.body || '', files, send: !!o.send }), 'utf8');
    fs.writeFileSync(path.join(dir, 'send.ps1'), '\ufeff' + PS, 'utf8');
  } catch (err) { return res({ ok: false, error: String(err) }); }
  const ch = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
    path.join(dir, 'send.ps1'), path.join(dir, 'job.json')], { windowsHide: true });
  let err = '';
  ch.stderr.on('data', (d) => { err += d; });
  const t = setTimeout(() => { ch.kill(); res({ ok: false, error: 'Zaman aşımı: Outlook yanıt vermedi.' }); }, 90000);
  ch.on('error', (x) => { clearTimeout(t); res({ ok: false, error: String(x) }); });
  ch.on('close', (c) => {
    clearTimeout(t);
    setTimeout(() => fs.rm(dir, { recursive: true, force: true }, () => {}), 60000);
    res(c === 0 ? { ok: true } : { ok: false, error: err.trim() || ('Çıkış kodu ' + c) });
  });
}));
