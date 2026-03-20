let colorData = null; // the full config_colores.json object
let ftpConfig  = {};
let loadedTpl = '';
let loadedLayout = '';
let activeGroup = null;
let editingIndex = null; // null = add, number = edit

const GROUP_META = {
    tapa:          { label: 'Tapa',          group: 'Tapa',          sub: 'Exterior superior de la valija',             texture2: null },
    base:          { label: 'Base',          group: 'Base',          sub: 'Exterior inferior de la valija',             texture2: null },
    detalles:      { label: 'Detalles',      group: 'Detalles',      sub: 'Logos, ruedas y detalles decorativos',       texture2: 'texture_ruedas' },
    separador:     { label: 'Separador',     group: 'Separador',     sub: 'Tela interior divisoria y cierre',           texture2: 'texture_cierre' },
    forreria:      { label: 'Forrería',      group: 'Forreria',      sub: 'Forro interior (opcional +$30.000)',         texture2: null },
    organizadores: { label: 'Organizadores', group: 'Organizadores', sub: 'Bolsos organizadores (opcional +$85.000)',   texture2: 'texture_chico' }
};

// ─── Init ─────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', async () => {
    // Load saved FTP config
    const saved = await window.api.getFtpConfig();
    if (saved.host)     document.getElementById('ftp-host').value = saved.host;
    if (saved.user)     document.getElementById('ftp-user').value = saved.user;
    if (saved.password) document.getElementById('ftp-pass').value = saved.password;
    if (saved.remoteTplPath) document.getElementById('ftp-tpl-path').value = saved.remoteTplPath;
    if (saved.remoteLayoutPath) document.getElementById('ftp-layout-path').value = saved.remoteLayoutPath;
    
    ftpConfig = saved;

    // FTP Toggle logic
    const ftpToggle = document.getElementById('ftp-toggle');
    const ftpFields = document.getElementById('ftp-fields-container');
    const ftpChevron = document.getElementById('ftp-chevron');
    
    ftpToggle.addEventListener('click', () => {
        const isHidden = ftpFields.style.display === 'none';
        ftpFields.style.display = isHidden ? 'block' : 'none';
        ftpChevron.textContent = isHidden ? '▲' : '▼';
    });

    document.getElementById('btn-connect').addEventListener('click', connectAndLoad);
    document.getElementById('btn-save').addEventListener('click', saveToFtp);
    document.getElementById('btn-reload').addEventListener('click', connectAndLoad);
    document.getElementById('btn-add-color').addEventListener('click', () => openModal(null));
    document.getElementById('btn-import-excel').addEventListener('click', handleImportExcel);
    document.getElementById('btn-export-excel').addEventListener('click', handleExportExcel);

});

// ─── FTP Actions ──────────────────────────────────────────────────────────────
function getFtpConfigFromUI() {
    return {
        host:       document.getElementById('ftp-host').value.trim(),
        user:       document.getElementById('ftp-user').value.trim(),
        password:   document.getElementById('ftp-pass').value,
        remotePath: '/static/config_colores.json', // Still used internally but hidden from UI
        remoteTplPath: document.getElementById('ftp-tpl-path').value.trim() || 'templates/page.tpl',
        remoteLayoutPath: document.getElementById('ftp-layout-path').value.trim() || 'layouts/layout.tpl'
    };
}

async function connectAndLoad() {
    const cfg = getFtpConfigFromUI();
    if (!cfg.host || !cfg.user) { showToast('Completá host y usuario', 'error'); return; }

    setStatus('Conectando...', '');
    setBtnLoading('btn-connect', true, '⬇ Conectar y Cargar');
    document.getElementById('btn-connect').disabled = true;

    const result = await window.api.ftpDownload(cfg);

    setBtnLoading('btn-connect', false, '⬇ Conectar y Cargar');
    document.getElementById('btn-connect').disabled = false;

    if (!result.ok) {
        setStatus('Error de conexión', 'error');
        let errMsg = result.error;
        if (result.dirList && result.dirList.length > 0) {
            errMsg += "\nArchivos encontrados en esa carpeta:\n" + result.dirList.join(", ");
        }
        showToast('❌ ' + errMsg, 'error');
        return;
    }

    ftpConfig = cfg;
    loadedTpl = result.tpl || '';
    loadedLayout = result.layout || '';

    // If JSON is empty or missing, try to parse from TPL
    if (!result.data || Object.keys(result.data).length === 0) {
        colorData = parseTplToColorData(loadedTpl, loadedLayout);
        showToast('Configuración extraída del page.tpl (JSON no encontrado)', 'info');
    } else {
        colorData = result.data;
    }
    
    setStatus('Conectado ✓ ' + cfg.host, 'connected');
    
    if (result.tplError || result.layoutError) {
        let warnMsg = "Carga parcial. Hubo problemas con plantillas:\n";
        if (result.tplError) warnMsg += "- Error en page.tpl: " + result.tplError + "\n";
        if (result.layoutError) warnMsg += "- Error en layout.tpl: " + result.layoutError;
        showToast('⚠️ ' + warnMsg, 'warning');
    } else {
        showToast('Plantillas cargadas correctamente', 'success');
    }

    renderGroupNav();
    if (colorData && Object.keys(colorData).length > 0) {
        showEditor(Object.keys(colorData)[0]);
    } else {
        showEmptyState("No se encontraron colores", "Aseguráte de que el page.tpl tenga las marcas de inyección.");
    }
}

async function saveToFtp() {
    if (!colorData) return;
    const cfg = getFtpConfigFromUI();
    setStatus('Guardando...', '');
    setBtnLoading('btn-save', true, '↑ Guardar en FTP');
    document.getElementById('btn-save').disabled = true;

    const newTpl = injectColorsIntoTpl(loadedTpl, colorData);
    let newLayout = null;
    if (loadedLayout) {
        newLayout = injectIdsIntoLayout(loadedLayout, colorData);
    }
    
    // We send an empty object for dataJson to skip JSON upload if we want to honor "borremos eso del json"
    // Or we can just send null. Let's send null.
    const result = await window.api.ftpUpload(cfg, null, newTpl, newLayout);

    setBtnLoading('btn-save', false, '↑ Guardar en FTP');
    document.getElementById('btn-save').disabled = false;

    if (!result.ok) {
        setStatus('Error al guardar', 'error');
        showToast('❌ ' + result.error, 'error');
        return;
    }

    setStatus('Conectado ✓ ' + cfg.host, 'connected');
    showToast('✅ Cambios guardados en TPLs correctamente', 'success');
}

// ─── UI Rendering ─────────────────────────────────────────────────────────────
function renderGroupNav() {
    const nav = document.getElementById('group-nav');
    nav.style.display = 'block';
    nav.innerHTML = '<div class="sidebar-label" style="padding: 8px 4px 6px;">Grupos</div>';

    Object.keys(GROUP_META).forEach(key => {
        if (!colorData[key]) return;
        const meta = GROUP_META[key];
        const count = colorData[key].length;
        const item = document.createElement('div');
        item.className = 'group-nav-item';
        item.dataset.group = key;
        item.innerHTML = `<span>${meta.label}</span><span class="count">${count}</span>`;
        item.addEventListener('click', () => showEditor(key));
        nav.appendChild(item);
    });

    document.getElementById('sidebar-bottom').style.display = 'block';
    document.getElementById('excel-section').style.display = 'block';
}

function showEditor(groupKey) {
    activeGroup = groupKey;

    // Update nav highlight
    document.querySelectorAll('.group-nav-item').forEach(el => {
        el.classList.toggle('active', el.dataset.group === groupKey);
    });

    const meta = GROUP_META[groupKey] || { label: groupKey, sub: '' };
    document.getElementById('editor-group-title').textContent = meta.label;
    document.getElementById('editor-group-sub').textContent = meta.sub;

    document.getElementById('empty-state').style.display = 'none';
    document.getElementById('editor-panel').style.display = 'flex';

    renderColorGrid(groupKey);
}

function renderColorGrid(groupKey) {
    const grid = document.getElementById('color-grid');
    grid.innerHTML = '';

    const colors = colorData[groupKey] || [];

    colors.forEach((c, i) => {
        const card = document.createElement('div');
        card.className = 'color-card';
        card.title = 'Clic para editar';
        card.addEventListener('click', (e) => {
            if (!e.target.classList.contains('color-card-del')) openModal(i);
        });

        // Swatch
        const swatch = document.createElement('div');
        swatch.className = 'color-card-swatch';
        if (c.type === 'pattern' && c.value && c.value.startsWith('http')) {
            swatch.style.backgroundImage = `url('${c.value}')`;
            swatch.style.backgroundSize = 'cover';
            swatch.style.backgroundPosition = 'center';
        } else {
            swatch.style.background = c.value || '#333';
        }

        // Delete btn
        const delBtn = document.createElement('button');
        delBtn.className = 'color-card-del';
        delBtn.textContent = '✕';
        delBtn.title = 'Eliminar';
        delBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            deleteColor(groupKey, i);
        });
        swatch.appendChild(delBtn);

        // Body
        const body = document.createElement('div');
        body.className = 'color-card-body';
        body.innerHTML = `
            <div class="color-card-name">${c.name}</div>
            <div class="color-card-meta">${c.type || 'solid'}</div>
            <div class="color-card-tn">${c.tn_id ? 'ID: ' + c.tn_id : '— sin ID —'}</div>
        `;

        card.appendChild(swatch);
        card.appendChild(body);
        grid.appendChild(card);
    });

    // Update count in nav
    const navItem = document.querySelector(`.group-nav-item[data-group="${groupKey}"] .count`);
    if (navItem) navItem.textContent = colors.length;
}

function deleteColor(groupKey, index) {
    colorData[groupKey].splice(index, 1);
    renderColorGrid(groupKey);
    showToast(`Color eliminado (no olvidés guardar)`, 'info');
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function openModal(index) {
    editingIndex = index;
    const isEdit = index !== null;
    const meta = GROUP_META[activeGroup];

    document.getElementById('modal-title').textContent = isEdit ? 'Editar color' : 'Agregar color';
    document.getElementById('btn-modal-save').textContent = isEdit ? 'Guardar cambios' : 'Agregar';

    // Show/hide secondary texture field based on group
    const tex2Field = document.getElementById('m-texture2-field');
    if (meta && meta.texture2) {
        tex2Field.style.display = 'block';
        document.getElementById('m-texture2-label').textContent =
            `URL Texture 2 (${meta.texture2})`;
    } else {
        tex2Field.style.display = 'none';
    }

    if (isEdit) {
        const c = colorData[activeGroup][index];
        document.getElementById('m-name').value    = c.name    || '';
        document.getElementById('m-type').value    = c.type    || 'solid';
        document.getElementById('m-value').value   = c.value   || '';
        document.getElementById('m-tnid').value    = c.tn_id   || '';
        document.getElementById('m-texture').value = c.texture || '';
        if (meta && meta.texture2) {
            document.getElementById('m-texture2').value = c[meta.texture2] || '';
        }
    } else {
        document.getElementById('m-name').value    = '';
        document.getElementById('m-type').value    = 'solid';
        document.getElementById('m-value').value   = '';
        document.getElementById('m-tnid').value    = '';
        document.getElementById('m-texture').value = '';
        document.getElementById('m-texture2').value = '';
    }

    updateSwatchType();
    updatePreview();
    document.getElementById('modal-overlay').classList.add('open');
    document.getElementById('m-name').focus();
}

function closeModal() {
    document.getElementById('modal-overlay').classList.remove('open');
    editingIndex = null;
}

// Close modal on overlay click
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('modal-overlay').addEventListener('click', (e) => {
        if (e.target === document.getElementById('modal-overlay')) closeModal();
    });
});

function updateSwatchType() {
    const type = document.getElementById('m-type').value;
    const label = document.getElementById('m-value-label');
    const input = document.getElementById('m-value');
    if (type === 'solid') {
        label.textContent = 'Color (hex)';
        input.placeholder = '#872734';
    } else {
        label.textContent = 'URL / Patrón';
        input.placeholder = 'https://... o linear-gradient(...)';
    }
    updatePreview();
}

function updatePreview() {
    const val = document.getElementById('m-value').value.trim();
    const preview = document.getElementById('m-preview');
    if (val.startsWith('http')) {
        preview.style.backgroundImage = `url('${val}')`;
        preview.style.background = '';
    } else {
        preview.style.backgroundImage = '';
        preview.style.background = val || '#333';
    }
}

function saveColorFromModal() {
    const name    = document.getElementById('m-name').value.trim();
    const type    = document.getElementById('m-type').value;
    const value   = document.getElementById('m-value').value.trim();
    const tn_id   = document.getElementById('m-tnid').value.trim();
    const texture = document.getElementById('m-texture').value.trim();
    const texture2 = document.getElementById('m-texture2').value.trim();
    const meta    = GROUP_META[activeGroup];

    if (!name) { showToast('El nombre es obligatorio', 'error'); return; }

    const colorObj = { name, type, value, tn_id, texture };
    if (meta && meta.texture2 && texture2) {
        colorObj[meta.texture2] = texture2;
    }

    if (editingIndex !== null) {
        colorData[activeGroup][editingIndex] = colorObj;
    } else {
        colorData[activeGroup].push(colorObj);
    }

    renderColorGrid(activeGroup);
    closeModal();
    showToast(editingIndex !== null ? 'Color actualizado' : 'Color agregado', 'success');
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function setStatus(text, cls) {
    const el = document.getElementById('status-text');
    el.textContent = text;
    el.className = 'titlebar-status' + (cls ? ' ' + cls : '');
}

function setBtnLoading(id, loading, label) {
    const btn = document.getElementById(id);
    const content = document.getElementById('connect-btn-content');
    if (loading) {
        content.innerHTML = '<div class="spinner" style="display:inline-block;"></div> Conectando...';
    } else {
        content.textContent = label;
    }
}

let toastTimeout;
function showToast(msg, type = 'info') {
    const toast = document.getElementById('toast');
    const msgEl = document.getElementById('toast-msg');
    toast.className = `toast ${type}`;
    msgEl.textContent = msg;
    toast.classList.add('show');
    clearTimeout(toastTimeout);
    toastTimeout = setTimeout(() => toast.classList.remove('show'), 3500);
}

// ESC closes modal
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
});


// --- GCS Upload ---------------------------------------------------------------
async function handleUploadClick(inputId) {
    const localPath = await window.api.selectFile();
    if (!localPath) return; // user canceled
    
    const btnId = 'btn-upload-' + inputId.replace('m-', '');
    const btn = document.getElementById(btnId);
    const originalText = btn.innerHTML;
    
    btn.innerHTML = '<div class="spinner" style="display:inline-block; vertical-align:middle; margin-right:4px;"></div> Subiendo...';
    btn.disabled = true;


    const result = await window.api.uploadTexture(localPath);
    
    btn.innerHTML = originalText;
    btn.disabled = false;

    if (!result.ok) {
        showToast('? Error subiendo a GCS: ' + result.error, 'error');
        return;
    }

    document.getElementById(inputId).value = result.url;
    showToast('? Imagen subida correctamente', 'success');
    if (inputId === 'm-value') updatePreview();
}

// ─── TPL Injection Generators ─────────────────────────────────────────────────
function buildProductIdsJS(data) {
    let lines = [];
    Object.keys(GROUP_META).forEach(jsonKey => {
        const cfg = GROUP_META[jsonKey];
        if (!data[jsonKey]) return;
        let groupIds = [];
        data[jsonKey].forEach(c => {
            if (c.tn_id) groupIds.push(`'${c.name.replace(/'/g, "\\'")}': '${c.tn_id}'`);
            else groupIds.push(`'${c.name.replace(/'/g, "\\'")}': ''`);
        });
        if (groupIds.length > 0) {
            lines.push(`            '${cfg.group}': { ${groupIds.join(', ')} }`);
        }
    });
    return '{\n' + lines.join(',\n') + '\n        }';
}

function buildTextureMapsJS(data) {
    const textureMap = {};
    const detallesMap = {};

    Object.keys(GROUP_META).forEach(jsonKey => {
        const cfg = GROUP_META[jsonKey];
        if (!data[jsonKey]) return;

        if (jsonKey === 'detalles') {
            data[jsonKey].forEach(c => {
                detallesMap[c.name] = { logo: c.texture, ruedas: c.texture_ruedas || c.texture };
            });
            return;
        }

        let matNameMap = {
            'Tapa': 'tapa_1',
            'Base': 'contratapa_1',
            'Separador': 'telas',
            'Forreria': 'forro',
            'Organizadores': 'bolso_grande'
        };
        let matName2Map = {
            'Separador': 'cierre',
            'Organizadores': 'bolso_chico'
        };

        const entry = { matName: matNameMap[cfg.group], files: {} };
        if (matName2Map[cfg.group]) { 
            entry.matName2 = matName2Map[cfg.group]; 
            entry.files2 = {}; 
        }

        data[jsonKey].forEach(c => {
            entry.files[c.name] = c.texture;
            if (entry.matName2) {
                if (jsonKey === 'separador') {
                    entry.files2[c.name] = c.texture_cierre || c.texture; 
                } else if (jsonKey === 'organizadores') {
                    entry.files2[c.name] = c.texture_chico || c.texture;
                }
            }
        });
        textureMap[cfg.group] = entry;
    });

    function toSingleQuotes(json) {
        return json.replace(/"/g, "'");           // Replace double with single quotes
    }
    
    return {
        textureMapStr: `        var textureMap = ${toSingleQuotes(JSON.stringify(textureMap, null, 4)).replace(/\n/g, '\n        ')};`,
        detallesMapStr: `            var detallesMap = ${toSingleQuotes(JSON.stringify(detallesMap, null, 4)).replace(/\n/g, '\n            ')};`
    };
}

function buildButtonsHTML(colorList, group) {
    if (!colorList) return '';
    return colorList.map(c => {
        let bgStyle = '';
        let bgAttr = '';
        if (c.type === 'pattern' && c.value && c.value.startsWith('http')) {
            bgAttr = ` data-bg="${c.value}"`;
        } else if (c.value) {
            bgStyle = ` style="background: ${c.value};"`;
        }
        const nameEscaped = c.name.replace(/"/g, '&quot;');
        const nameJsEscaped = c.name.replace(/'/g, "\\'");
        return `                        <button class="color-btn" data-group="${group}" data-color="${nameEscaped}"${bgStyle}${bgAttr} title="${nameEscaped}" onclick="cambiarColor(this,'${group}','${nameJsEscaped}')"></button>`;
    }).join('\n');
}

function injectColorsIntoTpl(tpl, data) {
    if (!tpl) return tpl;
    let newTpl = tpl;
    
    // 1. Inject Product IDs
    const productIdsStr = buildProductIdsJS(data);
    newTpl = newTpl.replace(/(window\.productIds\s*=\s*)\{[\s\S]*?\};/g, 
        (match, p1) => p1 + productIdsStr + ';');
        
    // 2. Inject Texture Maps
    const maps = buildTextureMapsJS(data);
    
    // Regex for textureMap
    newTpl = newTpl.replace(/(var\s+textureMap\s*=\s*)\{[\s\S]*?\};/g, 
        (match, p1) => {
            const objMatch = maps.textureMapStr.match(/var\s+textureMap\s*=\s*(\{[\s\S]*?\});/);
            return p1 + (objMatch ? objMatch[1] : '{}') + ';';
        });
        
    // Regex for detallesMap
    newTpl = newTpl.replace(/(var\s+detallesMap\s*=\s*)\{[\s\S]*?\};/g, 
        (match, p1) => {
            const objMatch = maps.detallesMapStr.match(/var\s+detallesMap\s*=\s*(\{[\s\S]*?\});/);
            return p1 + (objMatch ? objMatch[1] : '{}') + ';';
        });

    // 3. Inject HTML Buttons
    Object.keys(GROUP_META).forEach(jsonKey => {
        const groupLabel = GROUP_META[jsonKey].group;
        const buttonsHtml = buildButtonsHTML(data[jsonKey], groupLabel);
        
        // Target the color-list that follows the specific label ID
        const regex = new RegExp(`(id="text-label_${jsonKey}"[\\s\\S]*?<div class="color-list[^>]*">)([\\s\\S]*?)(<\\/div>)`, 'g');
        newTpl = newTpl.replace(regex, `$1\n${buttonsHtml}\n                    $3`);
    });

    return newTpl;
}

function injectIdsIntoLayout(layout, data) {
    const productIdsStr = buildProductIdsJS(data);
    return layout.replace(/(window\.productIds\s*=\s*)\{[\s\S]*?\};/g, 
        (match, p1) => p1 + productIdsStr + ';');
}

// ─── TPL PARSER ───────────────────────────────────────────────────────────────
function parseTplToColorData(tpl, layout) {
    const data = {};
    
    // Helper to extract JS object safely
    function extractJsObject(source, varName) {
        const regex = new RegExp(`(?:window\\.${varName}|var\\s+${varName}|let\\s+${varName})\\s*=\\s*({[\\s\\S]*?});`);
        const match = regex.exec(source);
        if (!match || !match[1]) {
            console.warn(`Could not find ${varName} assignment in source.`);
            return {};
        }
        
        try {
            return new Function(`return ${match[1].trim()}`)();
        } catch (e) {
            console.error(`Error parsing ${varName} from TPL:`, e, "Contenido extraído:", match[1]);
            return {};
        }
    }

    // 1. Extract Product IDs
    const fullSource = (layout || '') + '\n' + (tpl || '');
    const productIds = extractJsObject(fullSource, 'productIds');
 
    // 2. Extract Texture Maps
    const textureMap = extractJsObject(tpl, 'textureMap');
 
    // 3. Extract Detalles Map
    const detallesMap = extractJsObject(tpl, 'detallesMap');
 
    // 4. Extract Buttons per group
    Object.keys(GROUP_META).forEach(key => {
        // Target the color-list following the specific label ID
        const regex = new RegExp(`id="text-label_${key}"[\\s\\S]*?<div class="color-list[^>]*">([\\s\\S]*?)<\\/div>`, 'g');
        const match = regex.exec(tpl);
        
        if (match) {
            const buttonsHtml = match[1];
            const btnRegex = /<button[^>]*data-color="([^"]*)"[^>]*>/g;
            let btnMatch;
            data[key] = [];
            
            while ((btnMatch = btnRegex.exec(buttonsHtml)) !== null) {
                const fullTag = btnMatch[0];
                const name = btnMatch[1];
                
                let value = '';
                let type = 'solid';

                const bgMatch = fullTag.match(/data-bg="([^"]*)"/);
                const styleMatch = fullTag.match(/style="background:\s*([^;"]*);?"/);
                
                if (bgMatch) {
                    value = bgMatch[1];
                    type = 'pattern';
                } else if (styleMatch) {
                    value = styleMatch[1].trim();
                    if (value.includes('gradient') || value.includes('url')) {
                        type = 'pattern';
                    }
                }

                const groupLabel = GROUP_META[key].group;
                const colorObj = {
                    name: name,
                    type: type,
                    value: value,
                    tn_id: (productIds[groupLabel] && productIds[groupLabel][name]) || '',
                    texture: ''
                };

                // Extract texture
                if (key === 'detalles') {
                    if (detallesMap[name]) {
                        colorObj.texture = detallesMap[name].logo;
                        colorObj.texture_ruedas = detallesMap[name].ruedas;
                    }
                } else if (textureMap[groupLabel]) {
                    colorObj.texture = textureMap[groupLabel].files[name] || '';
                    if (textureMap[groupLabel].matName2) {
                        const tex2Key = GROUP_META[key].texture2;
                        colorObj[tex2Key] = (textureMap[groupLabel].files2 && textureMap[groupLabel].files2[name]) || '';
                    }
                }

                data[key].push(colorObj);
            }
        }
    });

    return data;
}

function showEmptyState(title, sub) {
    document.getElementById('editor-panel').style.display = 'none';
    const empty = document.getElementById('empty-state');
    empty.style.display = 'flex';
    empty.querySelector('.empty-state-title').textContent = title;
    empty.querySelector('.empty-state-sub').textContent = sub;
}

// ─── EXCEL HELPER ─────────────────────────────────────────────────────────────
async function handleImportExcel() {
    const localPath = await window.api.selectFile('excel');
    if (!localPath) return;

    const statusEl = document.getElementById('excel-status');
    statusEl.innerHTML = '<div class="spinner" style="display:inline-block;"></div> Importando...';
    statusEl.style.display = 'block';

    const result = await window.api.readExcel(localPath);
    if (!result.ok) {
        showToast("Error al leer Excel: " + result.error, "error");
        statusEl.textContent = "Error al leer archivo";
        return;
    }

    const data = result.data;
    if (!data || data.length < 2) {
        showToast("El archivo está vacío o no tiene el formato correcto.", "error");
        statusEl.textContent = "Formato incorrecto";
        return;
    }

    if (!confirm("¿Seguro que querés importar estos colores? Se reemplazarán los actuales.")) {
        statusEl.style.display = 'none';
        return;
    }

    // Process rows (headers are i=0)
    // Format: [Grupo, Nombre, Tipo, Valor, TiendaNube ID, Textura 1, Textura 2]
    const newColorData = {};
    Object.keys(GROUP_META).forEach(k => newColorData[k] = []);

    // Reverse map: Label -> key
    const labelToKey = {};
    Object.keys(GROUP_META).forEach(k => labelToKey[GROUP_META[k].label] = k);

    data.forEach((row, i) => {
        if (i === 0) return; // skip header
        if (row.length < 2) return;

        const [label, name, type, value, tn_id, tex1, tex2] = row;
        if (!name) return;

        const key = labelToKey[label] || label.toLowerCase();
        
        if (newColorData[key]) {
            const colorObj = {
                name: String(name || ''),
                type: String(type || 'solid'),
                value: String(value || '#ffffff'),
                tn_id: String(tn_id || ''),
                texture: String(tex1 || '')
            };
            const tex2Key = GROUP_META[key] ? GROUP_META[key].texture2 : null;
            if (tex2Key) {
                colorObj[tex2Key] = String(tex2 || '');
            }
            newColorData[key].push(colorObj);
        }
    });

    colorData = newColorData;
    renderGroupNav();
    const firstGroup = Object.keys(colorData).find(k => colorData[k] && colorData[k].length > 0);
    if (firstGroup) showEditor(firstGroup);
    showToast("✓ Importado correctamente desde " + localPath.split(/[\\/]/).pop(), "success");
    statusEl.textContent = "✓ Importado";
    setTimeout(() => statusEl.style.display = 'none', 2000);
}

async function handleExportExcel() {
    if (!colorData || Object.keys(colorData).length === 0) {
        showToast("No hay datos para exportar. Conectá al FTP primero.", "error");
        return;
    }

    const btn = document.getElementById('btn-export-excel');
    const originalText = btn.textContent;
    btn.textContent = '...';
    btn.disabled = true;

    // Prepare data for Excel (Array of Arrays)
    const rows = [
        ['Grupo', 'Nombre', 'Tipo', 'Valor (Color/URL)', 'TiendaNube ID', 'Textura 1', 'Textura 2']
    ];

    Object.keys(colorData).forEach(groupKey => {
        const groupLabel = GROUP_META[groupKey] ? GROUP_META[groupKey].label : groupKey;
        const tex2Key = GROUP_META[groupKey] ? GROUP_META[groupKey].texture2 : null;
        
        colorData[groupKey].forEach(c => {
            rows.push([
                groupLabel,
                c.name,
                c.type,
                c.value,
                c.tn_id || '',
                c.texture || '',
                tex2Key ? (c[tex2Key] || '') : ''
            ]);
        });
    });

    const result = await window.api.saveExcel(rows);
    btn.textContent = originalText;
    btn.disabled = false;

    if (!result.ok) {
        showToast("Error al exportar: " + result.error, "error");
    } else if (!result.canceled) {
        showToast("Excel exportado correctamente", "success");
    }
}

function renderExcelTable(data) {
    const table = document.getElementById('excel-table');
    table.innerHTML = '';
    
    if (!data || data.length === 0) {
        table.innerHTML = '<tr><td style="padding:20px; text-align:center;">El archivo está vacío.</td></tr>';
        return;
    }

    data.forEach((row, i) => {
        const tr = document.createElement('tr');
        if (i === 0) tr.style.background = '#222';
        tr.style.borderBottom = '1px solid #333';
        
        row.forEach(cell => {
            const td = document.createElement(i === 0 ? 'th' : 'td');
            td.style.padding = '8px 12px';
            td.style.textAlign = 'left';
            td.textContent = cell || '';
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });
}

function openExcelModal() {
    document.getElementById('excel-modal-overlay').classList.add('open');
}

function closeExcelModal() {
    document.getElementById('excel-modal-overlay').classList.remove('open');
}

// Close excel modal on ESC or click outside
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeExcelModal();
});
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('excel-modal-overlay').addEventListener('click', (e) => {
        if (e.target === document.getElementById('excel-modal-overlay')) closeExcelModal();
    });
});
