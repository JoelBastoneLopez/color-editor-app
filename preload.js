const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
    getFtpConfig:  ()             => ipcRenderer.invoke('get-ftp-config'),
    saveFtpConfig: (config)       => ipcRenderer.invoke('ftp-download', config), // Re-using download to save config as side effect
    saveGcsKey:    (localPath)    => ipcRenderer.invoke('save-gcs-key', localPath),
    googleLogin:   ()             => ipcRenderer.invoke('google-login'),
    ftpDownload:   (config)       => ipcRenderer.invoke('ftp-download', config),
    ftpUpload:     (config, data, tpl, layout) => ipcRenderer.invoke('ftp-upload', config, data, tpl, layout),
    selectFile:    (type)         => ipcRenderer.invoke('select-file', type),
    uploadTexture: (localPath)    => ipcRenderer.invoke('upload-texture', localPath),
    readExcel:     (localPath)    => ipcRenderer.invoke('read-excel', localPath),
    saveExcel:     (dataRows)     => ipcRenderer.invoke('save-excel', dataRows)
});
