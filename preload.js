const { contextBridge, ipcRenderer } = require('electron');
// Uygulamaya yalnızca e-posta gönderme köprüsü açılır (güvenlik için başka yetki yok).
contextBridge.exposeInMainWorld('native', {
  isDesktop: true,
  sendMail: (o) => ipcRenderer.invoke('mail', o)
});
