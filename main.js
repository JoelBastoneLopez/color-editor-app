const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const ftp = require('basic-ftp');
const XLSX = require('xlsx');
let CONFIG_FILE;

function createWindow() {
    const userDataPath = app.getPath('userData');
    if (!fs.existsSync(userDataPath)) fs.mkdirSync(userDataPath, { recursive: true });
    
    CONFIG_FILE = path.join(userDataPath, 'ftp_config.json');
    const win = new BrowserWindow({
        width: 1000,
        height: 700,
        icon: path.join(__dirname, 'assets/logo.png'),
        titleBarStyle: 'hidden',
        titleBarOverlay: {
            color: '#1a1a1a',
            symbolColor: '#f0f0f0',
            height: 52
        },
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        }
    });
    win.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('get-ftp-config', () => {
    let config = { 
        host: '', 
        user: '', 
        password: '', 
        remoteTplPath: 'templates/page.tpl', 
        remoteLayoutPath: 'layouts/layout.tpl' 
    };
    try {
        if (fs.existsSync(CONFIG_FILE)) {
            config = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf-8'));
        }
    } catch(e) {}
    
    return {
        ...config
    };
});


ipcMain.handle('ftp-download', async (event, config) => {
    const client = new ftp.Client();
    try {
        const configToSave = {
            host: config.host,
            user: config.user,
            password: config.password,
            remoteTplPath: config.remoteTplPath || 'templates/page.tpl',
            remoteLayoutPath: config.remoteLayoutPath || 'layouts/layout.tpl'
        };
        fs.writeFileSync(CONFIG_FILE, JSON.stringify(configToSave));

        await client.access({
            host: config.host,
            user: config.user,
            password: config.password,
            secure: true,
            secureOptions: { rejectUnauthorized: false }
        });


        // 2. Download page.tpl
        let tplContent = "";
        let tplError = null;
        if (config.remoteTplPath) {
            try {
                const tempTplPath = path.join(app.getPath('temp'), 'loaded_page.tpl');
                await client.downloadTo(tempTplPath, config.remoteTplPath);
                tplContent = fs.readFileSync(tempTplPath, 'utf-8');
            } catch (e) {
                tplError = e.message;
            }
        }

        // 3. Download layout.tpl
        let layoutContent = "";
        let layoutError = null;
        if (config.remoteLayoutPath) {
            try {
                const tempLayoutPath = path.join(app.getPath('temp'), 'loaded_layout.tpl');
                await client.downloadTo(tempLayoutPath, config.remoteLayoutPath);
                layoutContent = fs.readFileSync(tempLayoutPath, 'utf-8');
            } catch (e) {
                layoutError = e.message;
            }
        }

        return { 
            ok: true, 
            data: {}, 
            tpl: tplContent, 
            layout: layoutContent,
            tplError,
            layoutError
        };
    } catch (err) {
        // If JSON download or connection fails, try to list directory to help user
        let dirList = [];
        try {
            const parentDir = path.dirname(config.remotePath);
            const list = await client.list(parentDir);
            dirList = list.map(f => f.name);
        } catch(e) {}

        return { ok: false, error: err.message, dirList };
    } finally {
        client.close();
    }
});

ipcMain.handle('ftp-upload', async (event, config, dataJson, newTplContent, newLayoutContent) => {
    const client = new ftp.Client();
    try {
        await client.access({
            host: config.host,
            user: config.user,
            password: config.password,
            secure: true,
            secureOptions: { rejectUnauthorized: false }
        });


        // 2. Upload TPL
        if (config.remoteTplPath && newTplContent) {
            const tempTplPath = path.join(app.getPath('temp'), 'upload_page.tpl');
            fs.writeFileSync(tempTplPath, newTplContent, 'utf-8');
            await client.uploadFrom(tempTplPath, config.remoteTplPath);
        }

        // 3. Upload Layout
        if (config.remoteLayoutPath && newLayoutContent) {
            const tempLayoutPath = path.join(app.getPath('temp'), 'upload_layout.tpl');
            fs.writeFileSync(tempLayoutPath, newLayoutContent, 'utf-8');
            await client.uploadFrom(tempLayoutPath, config.remoteLayoutPath);
        }

        return { ok: true };
    } catch (err) {
        return { ok: false, error: err.message };
    } finally {
        client.close();
    }
});

ipcMain.handle('select-file', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Images', extensions: ['jpg', 'png', 'jpeg', 'webp'] }]
    });
    if (canceled || filePaths.length === 0) return null;
    return filePaths[0];
});

// Update CONFIG_FILE with gcsProxyUrl when called from get-ftp-config/ftp-download
ipcMain.handle('upload-texture', async (event, localPath) => {
    const client = new ftp.Client();
    try {
        if (!fs.existsSync(CONFIG_FILE)) {
            throw new Error("Primero configura los datos del FTP.");
        }
        const config = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf-8'));
        
        await client.access({
            host: config.host,
            user: config.user,
            password: config.password,
            secure: true,
            secureOptions: { rejectUnauthorized: false }
        });

        const fileName = path.basename(localPath).replace(/[^a-zA-Z0-9.-]/g, '_');
        const remoteTexturesDir = '/static/'; // Subimos directo a static para máxima compatibilidad con el CDN
        
        // Ensure textures directory exists
        await client.ensureDir(remoteTexturesDir);
        
        // Upload file
        await client.uploadFrom(localPath, remoteTexturesDir + fileName);

        // Pattern: https://dcdn-us.mitiendanube.com/stores/007/240/579/themes/rio/static/filename.png
        const publicUrl = `https://dcdn-us.mitiendanube.com/stores/007/240/579/themes/rio${remoteTexturesDir}${fileName}`;
        
        return { ok: true, url: publicUrl };
    } catch(err) {
        console.error("FTP Texture Upload Error:", err);
        return { ok: false, error: err.message };
    } finally {
        client.close();
    }
});

ipcMain.handle('read-excel', async (event, localPath) => {
    try {
        const workbook = XLSX.readFile(localPath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Array of arrays
        return { ok: true, data };
    } catch (e) {
        return { ok: false, error: e.message };
    }
});

ipcMain.handle('save-excel', async (event, dataRows) => {
    try {
        const { filePath } = await dialog.showSaveDialog({
            title: 'Exportar a Excel',
            defaultPath: path.join(app.getPath('documents'), 'Colores_TiendaNube.xlsx'),
            filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
        });
        if (!filePath) return { ok: true, canceled: true };

        const worksheet = XLSX.utils.aoa_to_sheet(dataRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Colores");
        XLSX.writeFile(workbook, filePath);
        
        return { ok: true };
    } catch (e) {
        return { ok: false, error: e.message };
    }
});
