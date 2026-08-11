const APP_VERSION = 'v1.2.0';

document.addEventListener('DOMContentLoaded', () => {
  const excelDropZone = document.getElementById('excelDropZone');
  const excelInput = document.getElementById('excelInput');
  const excelFileInfo = document.getElementById('excelFileInfo');
  
  const photoDirZone = document.getElementById('photoDirZone');
  const photoDirInput = document.getElementById('photoDirInput');
  const photoDirInfo = document.getElementById('photoDirInfo');
  
  const btnStartProcess = document.getElementById('btnStartProcess');
  const btnSaveToFolder = document.getElementById('btnSaveToFolder');
  const btnDownloadZip = document.getElementById('btnDownloadZip');
  const btnDownloadSample = document.getElementById('btnDownloadSample');
  
  const resultSection = document.getElementById('resultSection');
  const progressContainer = document.getElementById('progressContainer');
  const progressBar = document.getElementById('progressBar');
  const progressText = document.getElementById('progressText');
  const progressPercent = document.getElementById('progressPercent');
  const logList = document.getElementById('logList');
  
  const statTotalSku = document.getElementById('statTotalSku');
  const statMatchedImages = document.getElementById('statMatchedImages');
  const statMissingImages = document.getElementById('statMissingImages');
  
  const btnOpenConfig = document.getElementById('btnOpenConfig');
  const configModal = document.getElementById('configModal');
  const configModalClose = document.getElementById('configModalClose');
  const tabBtns = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');
  const btnSaveConfig = document.getElementById('btnSaveConfig');
  const btnResetConfig = document.getElementById('btnResetConfig');
  
  
  
  const guiConfigForm = document.getElementById('guiConfigForm');
  const guiCollectionForm = document.getElementById('guiCollectionForm');
  const btnAddCollection = document.getElementById('btnAddCollection');
  const guiColorForm = document.getElementById('guiColorForm');
  const btnAddColor = document.getElementById('btnAddColor');

  function renderGuiConfig(config) {
    let fixedHtml = '';
    for (const [k, v] of Object.entries(config.field_mappings.fixed || {})) {
      fixedHtml += `
        <div class="dynamic-row fixed-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${k}">
          <input type="text" class="val-input" placeholder="固定填寫內容" value="${v}">
          <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">刪除</button>
        </div>`;
    }

    let dynamicHtml = '';
    for (const [k, v] of Object.entries(config.field_mappings.dynamic || {})) {
      dynamicHtml += `
        <div class="dynamic-row dynamic-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${k}">
          <input type="text" class="val-input" placeholder="來源表對應欄位" value="${v}">
          <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">刪除</button>
        </div>`;
    }

    guiConfigForm.innerHTML = `
      <div class="form-group">
        <h4>來源表設定 (Source Settings)</h4>
        <div class="input-row"><label>來源表名稱:</label><input type="text" id="cfg_sheet_name" value="${config.source.sheet_name || ''}"></div>
        <div class="input-row"><label>標題列位於第幾列:</label><input type="number" id="cfg_header_row" value="${config.source.header_row || 3}"></div>
        <div class="input-row"><label>資料起始列:</label><input type="number" id="cfg_row_start" value="${config.source.row_start || 4}"></div>
        <div class="input-row"><label>篩選欄位(有Y才):</label><input type="text" id="cfg_filter_column" value="${config.source.filter_column || ''}"></div>
      </div>
      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 1px solid #e2e8f0; margin-bottom:12px; padding-bottom:8px;">
          <h4 style="border:none; margin:0; padding:0;">固定欄位對應 (Fixed)</h4>
          <button class="btn btn-outline btn-sm" onclick="addConfigFixedRow()">+ 新增</button>
        </div>
        <p style="font-size:0.85rem; color:#64748b; margin-bottom:10px;">不管來源資料為何，強制填入目標模板的固定值。</p>
        <div id="configFixedContainer">${fixedHtml}</div>
      </div>
      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 1px solid #e2e8f0; margin-bottom:12px; padding-bottom:8px;">
          <h4 style="border:none; margin:0; padding:0;">動態欄位對應 (Dynamic)</h4>
          <button class="btn btn-outline btn-sm" onclick="addConfigDynamicRow()">+ 新增</button>
        </div>
        <p style="font-size:0.85rem; color:#64748b; margin-bottom:10px;">將來源表的特定欄位資料，填入到目標模板的對應欄位中。</p>
        <div id="configDynamicContainer">${dynamicHtml}</div>
      </div>
    `;
  }

  window.addConfigFixedRow = function(k='', v='') {
    const div = document.createElement('div');
    div.className = 'dynamic-row fixed-field-row';
    div.innerHTML = `<input type="text" class="key-input" placeholder="目標模板欄位" value="${k}"><input type="text" class="val-input" placeholder="固定填寫內容" value="${v}"><button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">刪除</button>`;
    document.getElementById('configFixedContainer').appendChild(div);
  };

  window.addConfigDynamicRow = function(k='', v='') {
    const div = document.createElement('div');
    div.className = 'dynamic-row dynamic-field-row';
    div.innerHTML = `<input type="text" class="key-input" placeholder="目標模板欄位" value="${k}"><input type="text" class="val-input" placeholder="來源表對應欄位" value="${v}"><button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">刪除</button>`;
    document.getElementById('configDynamicContainer').appendChild(div);
  };

  function parseGuiConfig() {
    const fixed = {};
    document.querySelectorAll('#configFixedContainer .fixed-field-row').forEach(row => {
      const k = row.querySelector('.key-input').value.trim();
      const v = row.querySelector('.val-input').value.trim();
      if(k) fixed[k] = v;
    });

    const dynamic = {};
    document.querySelectorAll('#configDynamicContainer .dynamic-field-row').forEach(row => {
      const k = row.querySelector('.key-input').value.trim();
      const v = row.querySelector('.val-input').value.trim();
      if(k) dynamic[k] = v;
    });

    return {
      source: {
        file_path: "My Family.xlsx",
        sheet_name: document.getElementById('cfg_sheet_name').value.trim(),
        header_row: parseInt(document.getElementById('cfg_header_row').value) || 3,
        row_start: parseInt(document.getElementById('cfg_row_start').value) || 4,
        filter_column: document.getElementById('cfg_filter_column').value.trim()
      },
      field_mappings: {
        fixed: fixed,
        dynamic: dynamic
      }
    };
  }

  function renderGuiCollection(collectionAliases) {
    guiCollectionForm.innerHTML = '';
    for (const [k, vArr] of Object.entries(collectionAliases)) {
      addCollectionRow(k, vArr.join(', '));
    }
  }

  function addCollectionRow(key = '', aliases = '') {
    const row = document.createElement('div');
    row.className = 'dynamic-row';
    row.innerHTML = `
      <input type="text" class="key-input" placeholder="資料夾系列英文字眼 (例: HERMITAGE)" value="${key}">
      <input type="text" class="val-input" placeholder="品名中文系列名 (以逗號分隔，例: 隱士, 隱士系列)" value="${aliases}">
      <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">刪除</button>
    `;
    guiCollectionForm.appendChild(row);
  }

  if (btnAddCollection) {
    btnAddCollection.addEventListener('click', () => addCollectionRow());
  }

  function parseGuiCollection() {
    const aliases = {};
    guiCollectionForm.querySelectorAll('.dynamic-row').forEach(row => {
      const k = row.querySelector('.key-input').value.trim();
      const valStr = row.querySelector('.val-input').value;
      if (k) {
        aliases[k] = valStr.split(',').map(s => s.trim()).filter(s => s);
      }
    });
    return aliases;
  }

  function renderGuiColor(colorAliases) {
    guiColorForm.innerHTML = '';
    for (const [k, vArr] of Object.entries(colorAliases)) {
      addColorRow(k, vArr.join(', '));
    }
  }

  function addColorRow(key = '', aliases = '') {
    const row = document.createElement('div');
    row.className = 'dynamic-row';
    row.innerHTML = `
      <input type="text" class="key-input" placeholder="資料夾上的顏色字眼 (例: BLUE)" value="${key}">
      <input type="text" class="val-input" placeholder="Excel上的中文顏色名 (以逗號分隔)" value="${aliases}">
      <button class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">刪除</button>
    `;
    guiColorForm.appendChild(row);
  }

  btnAddColor.addEventListener('click', () => addColorRow());

  function parseGuiColor() {
    const aliases = {};
    guiColorForm.querySelectorAll('.dynamic-row').forEach(row => {
      const k = row.querySelector('.key-input').value.trim();
      const valStr = row.querySelector('.val-input').value;
      if (k) {
        aliases[k] = valStr.split(',').map(s => s.trim()).filter(s => s);
      }
    });
    return aliases;
  }

  const templateHarnessInput = document.getElementById('templateHarnessInput');
  const templateLeashInput = document.getElementById('templateLeashInput');
  const templateHarnessStatus = document.getElementById('templateHarnessStatus');
  const templateLeashStatus = document.getElementById('templateLeashStatus');
  const btnRescanCollection = document.getElementById('btnRescanCollection');
  const btnRescanColor = document.getElementById('btnRescanColor');

  let sourceExcelFile = null;
  let loadedWorkbook = null;
  let photoFilesArray = [];
  let processedResults = [];
  let harnessExcelBuffer = null;
  let leashExcelBuffer = null;
  let harnessCount = 0;
  let leashCount = 0;
  
  let currentConfig = window.AppConfig.get();
  let currentCollectionAliases = window.AppConfig.getCollectionAliases();
  let currentColorAliases = window.AppConfig.getColorAliases();

  // Load custom templates if available
  let customTemplates = { HARNESS: null, LEASH: null };
  try {
    const savedTemplates = localStorage.getItem('my_family_templates');
    if (savedTemplates) {
      customTemplates = JSON.parse(savedTemplates);
      if (customTemplates.HARNESS) templateHarnessStatus.textContent = "已套用自訂模板";
      if (customTemplates.LEASH) templateLeashStatus.textContent = "已套用自訂模板";
    }
  } catch (e) {}

  function logMessage(msg, type = 'info') {
    const li = document.createElement('li');
    li.className = `log-item ${type}`;
    li.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
    logList.appendChild(li);
    logList.scrollTop = logList.scrollHeight;
  }

  function setProgress(pct, text) {
    progressBar.style.width = `${pct}%`;
    progressPercent.textContent = `${pct}%`;
    if (text) progressText.textContent = text;
  }

  function checkReady() {
    btnStartProcess.disabled = !(sourceExcelFile && photoFilesArray.length > 0);
  }

  async function autoScanAndSyncMappings(showLog = true) {
    let scannedFolder = { collections: [], colors: [] };
    if (photoFilesArray && photoFilesArray.length > 0) {
      scannedFolder = window.MyFamilyProcessor.scanFolderStructure(photoFilesArray);
    }

    let excelData = { collections: {}, colors: {} };
    if (loadedWorkbook) {
      const headerRow = currentConfig.source.header_row || 3;
      const rowStart = currentConfig.source.row_start || 4;
      const filterColName = currentConfig.source.filter_column || '中文背標';
      excelData = window.MyFamilyProcessor.extractMappingsFromExcel(loadedWorkbook, headerRow, rowStart, filterColName);
    }

    if (scannedFolder.collections.length > 0 || Object.keys(excelData.collections).length > 0 ||
        scannedFolder.colors.length > 0 || Object.keys(excelData.colors).length > 0) {
      
      const merged = window.MyFamilyProcessor.mergeScannedAliases(
        currentCollectionAliases,
        currentColorAliases,
        scannedFolder,
        excelData
      );

      currentCollectionAliases = merged.collectionAliases;
      currentColorAliases = merged.colorAliases;

      // 如果設定視窗已開啟，即時更新其 DOM 內容
      renderGuiCollection(currentCollectionAliases);
      renderGuiColor(currentColorAliases);

      if (showLog) {
        let msgParts = [];
        if (scannedFolder.collections.length > 0 || scannedFolder.colors.length > 0) {
          msgParts.push(`照片目錄掃描出 ${scannedFolder.collections.length} 個系列、${scannedFolder.colors.length} 種顏色`);
        }
        if (Object.keys(excelData.collections).length > 0 || Object.keys(excelData.colors).length > 0) {
          msgParts.push(`已從 Excel 欄位自動對應 ${merged.stats.matchedExcelCount} 組中英名稱`);
        }
        if (msgParts.length > 0) {
          logMessage(`[自動對照] ${msgParts.join('，')}，已自動補全對照表！`, 'success');
        }
      }
    }
  }

  excelDropZone.addEventListener('click', () => excelInput.click());
  excelDropZone.addEventListener('dragover', (e) => { e.preventDefault(); excelDropZone.classList.add('dragover'); });
  excelDropZone.addEventListener('dragleave', () => excelDropZone.classList.remove('dragover'));
  excelDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    excelDropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleExcelFile(e.dataTransfer.files[0]);
  });
  excelInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleExcelFile(e.target.files[0]);
  });

  async function handleExcelFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      alert('請上傳 Excel 檔案 (.xlsx 或 .xls)');
      return;
    }
    sourceExcelFile = file;
    excelDropZone.classList.add('has-file');
    excelFileInfo.textContent = `已選擇: ${file.name}`;
    excelFileInfo.classList.remove('hidden');
    logMessage(`已載入來源 Excel: ${file.name}`);
    checkReady();

    try {
      const arrayBuffer = await file.arrayBuffer();
      loadedWorkbook = await XlsxPopulate.fromDataAsync(arrayBuffer);
      await autoScanAndSyncMappings(true);
    } catch (err) {
      console.warn('解析 Excel 欄位對照失敗:', err);
    }
  }

  photoDirZone.addEventListener('click', () => photoDirInput.click());
  photoDirInput.addEventListener('change', async (e) => {
    if (e.target.files.length) {
      photoFilesArray = Array.from(e.target.files);
      photoDirZone.classList.add('has-file');
      photoDirInfo.textContent = `已選取資料夾，共包含 ${photoFilesArray.length} 個檔案`;
      photoDirInfo.classList.remove('hidden');
      logMessage(`已載入照片資料夾，包含 ${photoFilesArray.length} 個檔案`);
      checkReady();

      try {
        await autoScanAndSyncMappings(true);
      } catch (err) {
        console.warn('照片資料夾自動掃描失敗:', err);
      }
    }
  });

  if (btnRescanCollection) {
    btnRescanCollection.addEventListener('click', async () => {
      await autoScanAndSyncMappings(true);
      alert('已依目前載入之檔案重新掃描並更新系列對照表！');
    });
  }

  if (btnRescanColor) {
    btnRescanColor.addEventListener('click', async () => {
      await autoScanAndSyncMappings(true);
      alert('已依目前載入之檔案重新掃描並更新顏色別名對照表！');
    });
  }

  btnOpenConfig.addEventListener('click', () => {
    renderGuiConfig(currentConfig);
    renderGuiCollection(currentCollectionAliases);
    renderGuiColor(currentColorAliases);
    configModal.classList.remove('hidden');
  });
  
  configModalClose.addEventListener('click', () => configModal.classList.add('hidden'));
  
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      tabContents.forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.target).classList.add('active');
    });
  });

  function readAsBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(',')[1]);
      reader.onerror = error => reject(error);
      reader.readAsDataURL(file);
    });
  }

  btnSaveConfig.addEventListener('click', async () => {
    try {
      currentConfig = parseGuiConfig();
      currentCollectionAliases = parseGuiCollection();
      currentColorAliases = parseGuiColor();
      
      localStorage.setItem('my_family_config', JSON.stringify(currentConfig));
      localStorage.setItem('my_family_collection_aliases', JSON.stringify(currentCollectionAliases));
      localStorage.setItem('my_family_color_aliases', JSON.stringify(currentColorAliases));
      
      let updatedTmpl = false;
      if (templateHarnessInput.files.length > 0) {
        customTemplates.HARNESS = await readAsBase64(templateHarnessInput.files[0]);
        templateHarnessStatus.textContent = "已套用自訂模板";
        updatedTmpl = true;
      }
      if (templateLeashInput.files.length > 0) {
        customTemplates.LEASH = await readAsBase64(templateLeashInput.files[0]);
        templateLeashStatus.textContent = "已套用自訂模板";
        updatedTmpl = true;
      }
      if (updatedTmpl) {
        localStorage.setItem('my_family_templates', JSON.stringify(customTemplates));
      }

      alert('設定已儲存！');
      configModal.classList.add('hidden');
    } catch (e) {
      alert('JSON 格式錯誤，請檢查後再儲存。\n錯誤訊息: ' + e.message);
    }
  });

  btnResetConfig.addEventListener('click', () => {
    if (confirm('確定要還原為系統預設值嗎？')) {
      localStorage.removeItem('my_family_config');
      localStorage.removeItem('my_family_collection_aliases');
      localStorage.removeItem('my_family_color_aliases');
      localStorage.removeItem('my_family_templates');
      
      currentConfig = window.AppConfig.getDefaultConfig();
      currentCollectionAliases = window.AppConfig.getDefaultCollectionAliases();
      currentColorAliases = window.AppConfig.getDefaultColorAliases();
      customTemplates = { HARNESS: null, LEASH: null };
      
      renderGuiConfig(currentConfig);
      renderGuiCollection(currentCollectionAliases);
      renderGuiColor(currentColorAliases);
      
      templateHarnessStatus.textContent = "已套用預設模板";
      templateLeashStatus.textContent = "已套用預設模板";
      templateHarnessInput.value = "";
      templateLeashInput.value = "";
      
      alert('已還原為預設值。');
    }
  });

  btnDownloadSample.addEventListener('click', async () => {
    const zip = new JSZip();
    zip.file("README.txt", "請將來源 Excel 拖放至網頁上，並點擊按鈕選擇此 Photo 資料夾。\n");
    const photoFolder = zip.folder("Photo");
    const sampleProduct = photoFolder.folder("AMALFI DOG COLLAR 葡萄紫");
    
    const canvas = document.createElement('canvas');
    canvas.width = 600; canvas.height = 600;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#CCCCCC'; ctx.fillRect(0,0,600,600);
    ctx.fillStyle = '#000000'; ctx.font = '40px Arial'; ctx.fillText('Sample Image', 180, 300);
    
    const dummyBlob = await new Promise(r => canvas.toBlob(r, 'image/jpeg'));
    sampleProduct.file("主圖.jpg", dummyBlob);
    sampleProduct.file("情境圖1.jpg", dummyBlob);
    
    const content = await zip.generateAsync({type:"blob"});
    saveAs(content, "MyFamily_Sample_Folders.zip");
  });

  let filesToExport = {
    '胸背帶': { images: new Map(), back_labels: new Map() },
    '項圈牽繩': { images: new Map(), back_labels: new Map() }
  };

  function sanitizeFilename(name) {
    if (!name) return '';
    let s = String(name).replace(/[\r\n]+/g, ' ').trim();
    s = s.replace(/[\/\\*?:"<>|]/g, '_').trim();
    s = s.replace(/_+/g, '_').replace(/^_+|_+$/g, '');
    return s;
  }

  function getBaseProductName(prodName, size) {
    if (!prodName) return '';
    let base = String(prodName).trim();
    base = base.replace(/Ｍ/g, 'M').replace(/Ｓ/g, 'S').replace(/Ｌ/g, 'L').replace(/Ｘ/g, 'X');
    if (size && typeof size === 'string' && size.trim()) {
      let s = size.trim().replace(/Ｍ/g, 'M').replace(/Ｓ/g, 'S').replace(/Ｌ/g, 'L').replace(/Ｘ/g, 'X');
      let escaped = s.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      let pattern = new RegExp('(?:\\s*)' + escaped + '(?:號)?$', 'i');
      base = base.replace(pattern, '').trim();
    }
    base = base.replace(/(?:\s*)(?:2XS|3XS|4XS|5XS|XXS|XS|S|M|L|XL|XXL|2XL|3XL|4XL|5XL|\d{2})號?$/i, '').trim();
    return base;
  }

  function formatBarcode(val) {
    if (val === null || val === undefined || val === '') return '';
    let str = (typeof val === 'number') ? val.toLocaleString('fullwide', { useGrouping: false }) : String(val).trim();
    if (str.toUpperCase().includes('E')) {
      const num = Number(str);
      if (!isNaN(num)) {
        str = num.toLocaleString('fullwide', { useGrouping: false });
      }
    }
    if (str.includes('.')) {
      const parts = str.split('.');
      if (parts[1] === '0' || parts[1] === '00' || parts[1] === '000') {
        str = parts[0];
      } else {
        const num = Number(str);
        if (!isNaN(num)) {
          str = Math.round(num).toString();
        }
      }
    }
    return str;
  }

  function getCellValue(cell) {
    let v = cell.value();
    if (v === null || v === undefined) return '';
    if (typeof v === 'object' && v !== null) {
      if (typeof v.text === 'function') return v.text();
      return v.toString();
    }
    return v;
  }

  function b64ToArrayBuffer(base64) {
    const binaryString = window.atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }

  btnStartProcess.addEventListener('click', async () => {
    btnStartProcess.disabled = true;
    btnSaveToFolder.disabled = true;
    btnDownloadZip.disabled = true;
    resultSection.classList.remove('hidden');
    progressContainer.classList.remove('hidden');
    logList.innerHTML = '';
    processedResults = [];
    harnessExcelBuffer = null;
    leashExcelBuffer = null;
    harnessCount = 0;
    leashCount = 0;
    filesToExport = {
      '胸背帶': { images: new Map(), back_labels: new Map() },
      '項圈牽繩': { images: new Map(), back_labels: new Map() }
    };
    
    statTotalSku.textContent = '0';
    statMatchedImages.textContent = '0';
    statMissingImages.textContent = '0';
    
    try {
      logMessage('開始解析來源 Excel...');
      const arrayBuffer = await sourceExcelFile.arrayBuffer();
      const workbook = await XlsxPopulate.fromDataAsync(arrayBuffer);
      
      const sheetNames = ['MYFAMILY', 'MY FAMILY', 'My Family', 'Sheet1'];
      let ws = null;
      for (const name of sheetNames) {
        ws = workbook.sheet(name);
        if (ws) break;
      }
      if (!ws) ws = workbook.sheet(0);
      if (!ws) throw new Error('找不到任何工作表，請確認 Excel 檔案格式');
      logMessage(`使用工作表: ${ws.name()}`);

      const headerRow = currentConfig.source.header_row || 3;
      const rowStart = currentConfig.source.row_start || 4;
      const headers = {};
      
      const usedRange = ws.usedRange();
      if (!usedRange) throw new Error('來源工作表似乎是空的');
      const maxCol = usedRange.endCell().columnNumber();
      const totalRows = usedRange.endCell().rowNumber();

      for (let c = 1; c <= maxCol; c++) {
        const val = ws.cell(headerRow, c).value();
        if (val) {
          const key = val.toString().trim();
          if (!(key in headers)) headers[key] = c;
        }
      }
      logMessage(`已讀取表頭欄位: ${Object.keys(headers).length} 欄`);

      const processor = new window.MyFamilyProcessor(
        photoFilesArray,
        arrayBuffer, // not used directly in processor anymore, but kept for signature
        currentColorAliases,
        window.AppConfig.zhColorMap,
        window.AppConfig.targetCombos,
        currentCollectionAliases
      );

      // Load Templates (Custom or Built-in) using XlsxPopulate to preserve EXACT formatting (dropdowns, validations)
      const harnessB64 = customTemplates.HARNESS || window.MyFamilyTemplates?.HARNESS;
      const leashB64 = customTemplates.LEASH || window.MyFamilyTemplates?.LEASH;
      
      let outWbHarness = null;
      let outWsHarness = null;
      if (harnessB64) {
        outWbHarness = await XlsxPopulate.fromDataAsync(b64ToArrayBuffer(harnessB64));
        outWsHarness = outWbHarness.sheets().find(s => s.name().startsWith('QF_')) || outWbHarness.sheet(0);
      } else {
        outWbHarness = await XlsxPopulate.fromBlankAsync();
        outWsHarness = outWbHarness.sheet(0).name('QF_8051_胸背帶');
      }

      let outWbLeash = null;
      let outWsLeash = null;
      if (leashB64) {
        outWbLeash = await XlsxPopulate.fromDataAsync(b64ToArrayBuffer(leashB64));
        outWsLeash = outWbLeash.sheets().find(s => s.name().startsWith('QF_')) || outWbLeash.sheet(0);
      } else {
        outWbLeash = await XlsxPopulate.fromBlankAsync();
        outWsLeash = outWbLeash.sheet(0).name('QF_8051_項圈牽繩');
      }
      
      const harnessMap = {};
      const harnessMaxCol = outWsHarness.usedRange() ? outWsHarness.usedRange().endCell().columnNumber() : 60;
      for (let c = 1; c <= harnessMaxCol; c++) {
          const val = outWsHarness.cell(5, c).value();
          if (val) {
            const s = val.toString();
            harnessMap[s] = c;
            harnessMap[s.trim()] = c;
          }
      }

      const leashMap = {};
      const leashMaxCol = outWsLeash.usedRange() ? outWsLeash.usedRange().endCell().columnNumber() : 60;
      for (let c = 1; c <= leashMaxCol; c++) {
          const val = outWsLeash.cell(5, c).value();
          if (val) {
            const s = val.toString();
            leashMap[s] = c;
            leashMap[s.trim()] = c;
          }
      }
      
      let harnessRowIdx = 9;
      let leashRowIdx = 9;
      
      let totalCount = 0;
      let matchCount = 0;
      let missCount = 0;

      const filterColName = currentConfig.source.filter_column || '中文背標';
      const filterColIdx = headers[filterColName];
      
      const imageCounters = { '胸背帶': 1, '項圈牽繩': 1 };
      const prevBaseNames = { '胸背帶': '', '項圈牽繩': '' };
      const lastMainImgNames = { '胸背帶': '', '項圈牽繩': '' };
      const lastSc1ImgNames = { '胸背帶': '', '項圈牽繩': '' };
      const lastSc2ImgNames = { '胸背帶': '', '項圈牽繩': '' };
      const lastChartImgNames = { '胸背帶': '', '項圈牽繩': '' };
      
      for (let r = rowStart; r <= totalRows; r++) {
        if (filterColIdx) {
          const filterVal = (getCellValue(ws.cell(r, filterColIdx)) || '').toString().trim().toUpperCase();
          if (filterVal !== 'Y' && filterVal !== 'YES' && filterVal !== 'Ｙ') continue;
        }
        
        const firstCell = ws.cell(r, 1);
        const fill = firstCell.style('fill');
        if (fill && fill.type === 'solid' && fill.color && typeof fill.color === 'string') {
          if (fill.color.toUpperCase().includes('FAD9D6')) continue;
        }

        const zhName = headers['中文品名'] ? (getCellValue(ws.cell(r, headers['中文品名'])) || '').toString().trim() : '';
        if (!zhName) continue;

        const skuVal = headers['SKU'] ? getCellValue(ws.cell(r, headers['SKU'])) : '';
        const skuStr = formatBarcode(skuVal);
        const rawCollection = headers['COLLECTION'] ? (getCellValue(ws.cell(r, headers['COLLECTION'])) || '').toString().trim() : '';
        const rawType = headers['TYPE'] ? (getCellValue(ws.cell(r, headers['TYPE'])) || '').toString().trim() : '';
        const rawColor = headers['COLOR'] ? (getCellValue(ws.cell(r, headers['COLOR'])) || '').toString().trim() : '';
        const rawSize = headers['SIZE'] ? (getCellValue(ws.cell(r, headers['SIZE'])) || '').toString().trim() : '';
        
        const parsed = processor.parseChineseName(zhName, rawSize);
        const rawHints = { collection: rawCollection, type: rawType, color: rawColor, size: rawSize, sku: skuStr };
        const targetInfo = processor.getTargetTemplateAndCategory(zhName, parsed.type || rawType);
        totalCount++;

        let imgs = { main: null, sc1: null, sc2: null, chart: null, label: null };
        if (parsed.success) {
            imgs = processor.getLocalImagesForProduct(parsed, rawHints);
        } else {
            logMessage(`[警告] 品名格式不符: ${zhName}，將以既有線索嘗試配對`, 'warning');
            imgs = processor.getLocalImagesForProduct(parsed, rawHints);
        }

        const targetSubFolder = (targetInfo.target_subfolder === '胸背帶') ? '胸背帶' : '項圈牽繩';
        const baseProdName = getBaseProductName(zhName, rawSize || parsed.size);
        const isSameProduct = (baseProdName !== '' && baseProdName === prevBaseNames[targetSubFolder]);

        let mainImgName = '';
        let sc1ImgName = '';
        let sc2ImgName = '';
        let chartImgName = '';

        if (isSameProduct) {
          mainImgName = lastMainImgNames[targetSubFolder] || '';
          sc1ImgName = lastSc1ImgNames[targetSubFolder] || '';
          sc2ImgName = lastSc2ImgNames[targetSubFolder] || '';
          chartImgName = lastChartImgNames[targetSubFolder] || '';
        } else {
          const cleanName = sanitizeFilename(baseProdName || zhName);

          if (imgs.main) {
            try {
              const blob = await window.ImageUtils.ensureMinWidth(await window.ImageUtils.loadImage(imgs.main));
              mainImgName = `${cleanName}主圖.jpg`;
              filesToExport[targetSubFolder].images.set(mainImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 主圖處理失敗: ${imgs.main.name}`, 'warning');
            }
          }
          lastMainImgNames[targetSubFolder] = mainImgName;

          if (imgs.sc1) {
            try {
              const blob = await window.ImageUtils.ensureMinWidth(await window.ImageUtils.loadImage(imgs.sc1));
              sc1ImgName = `${imageCounters[targetSubFolder]}.jpg`;
              imageCounters[targetSubFolder]++;
              filesToExport[targetSubFolder].images.set(sc1ImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 情境圖 1 處理失敗: ${imgs.sc1.name}`, 'warning');
            }
          }
          lastSc1ImgNames[targetSubFolder] = sc1ImgName;

          if (imgs.sc2) {
            try {
              const blob = await window.ImageUtils.ensureMinWidth(await window.ImageUtils.loadImage(imgs.sc2));
              sc2ImgName = `${imageCounters[targetSubFolder]}.jpg`;
              imageCounters[targetSubFolder]++;
              filesToExport[targetSubFolder].images.set(sc2ImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 情境圖 2 處理失敗: ${imgs.sc2.name}`, 'warning');
            }
          }
          lastSc2ImgNames[targetSubFolder] = sc2ImgName;

          if (imgs.chart) {
            try {
              const blob = await window.ImageUtils.ensureMinWidth(await window.ImageUtils.loadImage(imgs.chart));
              chartImgName = `尺寸規格表_${cleanName}.jpg`;
              filesToExport[targetSubFolder].images.set(chartImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 尺寸圖處理失敗: ${imgs.chart.name}`, 'warning');
            }
          }
          if (!chartImgName) {
            chartImgName = sc1ImgName || sc2ImgName || mainImgName;
          }
          lastChartImgNames[targetSubFolder] = chartImgName;

          prevBaseNames[targetSubFolder] = baseProdName;
        }

        let labelImgName = '';
        if (imgs.label) {
          try {
            const blob = await window.ImageUtils.ensureMinWidth(await window.ImageUtils.loadImage(imgs.label));
            const cleanProdName = sanitizeFilename(zhName);
            const szSuffix = (rawSize || parsed.size) ? `_${String(rawSize || parsed.size).trim()}` : '';
            labelImgName = `背標_${cleanProdName}${szSuffix}.jpg`;
            filesToExport[targetSubFolder].back_labels.set(labelImgName, blob);
          } catch(imgErr) {
            logMessage(`[警告] 背標處理失敗: ${imgs.label.name}`, 'warning');
          }
        } else {
          logMessage(`[提示] 找不到背標圖片: SKU=${skuStr}`, 'info');
        }

        if (mainImgName || imgs.main) {
          matchCount++;
        } else {
          missCount++;
          logMessage(`[警告] 找不到主要圖片: SKU=${skuStr}, 品名=${zhName}`, 'warning');
        }

        let targetWs, targetRowIdx, tmplMap;
        if (targetSubFolder === '胸背帶') {
          targetWs = outWsHarness;
          targetRowIdx = harnessRowIdx++;
          tmplMap = harnessMap;
          harnessCount++;
        } else {
          targetWs = outWsLeash;
          targetRowIdx = leashRowIdx++;
          tmplMap = leashMap;
          leashCount++;
        }
        
        const getSourceHeaderIdx = (colName) => {
          if (!colName) return undefined;
          if (headers[colName]) return headers[colName];
          if (headers[colName.trim()]) return headers[colName.trim()];
          const cleanK = colName.replace(/\s+/g, '');
          for (const [hk, idx] of Object.entries(headers)) {
            if (hk.replace(/\s+/g, '') === cleanK) return idx;
          }
          return undefined;
        };

        const setVal = (colName, val) => {
          const colIdx = tmplMap[colName] || tmplMap[colName.trim()];
          if (colIdx) {
            const cell = targetWs.cell(targetRowIdx, colIdx);
            const isBarcodeField = colName.includes('條碼') || colName.includes('EAN') || colName.includes('SKU') || colName.includes('GTIN') || colName.includes('Part Number') || colName.includes('Global Trade Item Number');
            if (isBarcodeField) {
              const barcodeStr = formatBarcode(val);
              cell.value(barcodeStr);
              try {
                cell.style('numberFormat', '@');
              } catch(e) {}
            } else {
              cell.value(val);
            }
            if (!parsed.success || ((!mainImgName && !imgs.main) && ['商品名稱', '商品正面(主要圖片）'].includes(colName))) {
              cell.style('fill', 'ffff0000');
            }
          }
        };

        setVal('細分商品種類', targetInfo.category);
        for (const [tKey, fVal] of Object.entries(currentConfig.field_mappings.fixed || {})) {
            if (['細分商品種類'].includes(tKey)) continue;
            setVal(tKey, fVal);
        }

        if (parsed.brand) setVal('品牌', parsed.brand);
        if (parsed.collection) setVal('系列', parsed.collection);
        if (parsed.color) setVal('顏色', parsed.color);
        if (parsed.size) setVal('尺寸', parsed.size);
        setVal('商品名稱', zhName);

        for (const [tKey, sCol] of Object.entries(currentConfig.field_mappings.dynamic || {})) {
            if (['細分商品種類'].includes(tKey)) continue;
            const srcIdx = getSourceHeaderIdx(sCol);
            if (srcIdx) {
                const rawVal = getCellValue(ws.cell(r, srcIdx));
                let writeVal;
                const isBarcode = tKey.includes('條碼') || tKey.includes('EAN') || tKey.includes('SKU') || sCol.includes('條碼') || sCol.includes('EAN') || sCol.includes('SKU');
                
                if (isBarcode) {
                  writeVal = formatBarcode(rawVal);
                } else {
                  writeVal = (typeof rawVal === 'number') ? rawVal : (rawVal || '').toString().trim();
                }

                if (tKey.includes('包裝尺寸') && writeVal) {
                  let s = writeVal.toString().trim();
                  const isCm = s.toLowerCase().includes('cm');
                  let cleanS = s.replace(/(mm|cm)/gi, '').trim().replace(/x/gi, '*').replace(/×/g, '*');
                  const parts = cleanS.split('*').map(p => p.trim()).filter(Boolean);
                  if (parts.length === 3) {
                    if (isCm) {
                      writeVal = parts.map(p => Math.round(parseFloat(p) * 10)).join('*');
                    } else {
                      writeVal = parts.map(p => Math.round(parseFloat(p))).join('*');
                    }
                  }
                }

                if (tKey.includes('包裝重量') && writeVal) {
                  try {
                    writeVal = Math.round(parseFloat(writeVal));
                  } catch (e) {}
                }

                setVal(tKey, writeVal);
            }
        }

        if (mainImgName) setVal('商品正面(主要圖片）', mainImgName);
        if (sc1ImgName) setVal('商品側拍或情境圖 1\n(消費者可見圖片）', sc1ImgName);
        if (sc2ImgName) setVal('商品側拍或情境圖 2\n(消費者可見圖片）', sc2ImgName);
        if (chartImgName) {
            for (let k in tmplMap) {
                if (k.includes('詳細說明圖集') || k.includes('尺寸圖') || k.includes('尺寸規格')) {
                    setVal(k, chartImgName);
                }
            }
        }
        if (labelImgName) {
            for (let k in tmplMap) {
                if (k.includes('背標')) {
                    setVal(k, labelImgName);
                }
            }
        }

        processedResults.push({
          skuStr, zhName, subFolder: targetSubFolder
        });

        if (r % 5 === 0 || r === totalRows) {
          const pct = Math.min(Math.round((r / totalRows) * 100), 100);
          setProgress(pct);
          statTotalSku.textContent = totalCount;
          statMatchedImages.textContent = matchCount;
          statMissingImages.textContent = missCount;
          await new Promise(resolve => setTimeout(resolve, 0));
        }
      }

      setProgress(100, '處理完成！請選擇匯出方式。');
      statTotalSku.textContent = totalCount;
      statMatchedImages.textContent = matchCount;
      statMissingImages.textContent = missCount;

      if (harnessCount > 0) {
        harnessExcelBuffer = await outWbHarness.outputAsync();
      } else {
        harnessExcelBuffer = null;
      }

      if (leashCount > 0) {
        leashExcelBuffer = await outWbLeash.outputAsync();
      } else {
        leashExcelBuffer = null;
      }

      if (totalCount === 0) {
        btnSaveToFolder.disabled = true;
        btnDownloadZip.disabled = true;
        logMessage('處理完成，但未找到任何符合篩選條件的商品資料。', 'warning');
      } else {
        btnSaveToFolder.disabled = false;
        btnDownloadZip.disabled = false;
        const typesDesc = [];
        if (harnessCount > 0) typesDesc.push(`胸背帶: ${harnessCount} 筆`);
        if (leashCount > 0) typesDesc.push(`項圈牽繩: ${leashCount} 筆`);
        logMessage(`全部處理完成！共 ${totalCount} 筆 SKU（${typesDesc.join('，')}），配對成功 ${matchCount} 筆，缺圖 ${missCount} 筆`, 'success');
      }

    } catch(e) {
      console.error(e);
      logMessage(`處理發生錯誤: ${e.message}`, 'error');
      progressText.textContent = '處理失敗';
    } finally {
      btnStartProcess.disabled = false;
    }
  });

  btnSaveToFolder.addEventListener('click', async () => {
    if (processedResults.length === 0) return;

    if (!('showDirectoryPicker' in window)) {
      alert('您的瀏覽器不支援「直接存入資料夾」功能（建議使用 Chrome 或 Edge 瀏覽器）。\n系統將自動為您切換為「打包下載 ZIP」！');
      btnDownloadZip.click();
      return;
    }

    let dirHandle;
    try {
      dirHandle = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'downloads' });
    } catch (err) {
      if (err.name === 'AbortError') return;
      console.error('選取資料夾失敗:', err);
      alert('無法存取該資料夾，請確認權限或改用「打包下載 ZIP」。');
      return;
    }

    progressContainer.classList.remove('hidden');
    setProgress(0, '正在寫入檔案至所選資料夾...');

    try {
      const savedCategories = [];
      let totalFiles = 0;

      // 1. 胸背帶 (僅在有資料時輸出)
      const hData = filesToExport['胸背帶'];
      if (harnessCount > 0 || harnessExcelBuffer || hData.images.size > 0 || hData.back_labels.size > 0) {
        savedCategories.push('胸背帶');
        const harnessDir = await dirHandle.getDirectoryHandle('胸背帶', { create: true });
        
        if (harnessExcelBuffer) {
          const fh = await harnessDir.getFileHandle('商品報價單_胸背帶_auto_generate.xlsx', { create: true });
          const w = await fh.createWritable();
          await w.write(harnessExcelBuffer);
          await w.close();
        }

        if (hData.images.size > 0) {
          const harnessImgDir = await harnessDir.getDirectoryHandle('images', { create: true });
          for (const [filename, blob] of hData.images.entries()) {
            await writeBlob(harnessImgDir, filename, blob);
          }
        }
        if (hData.back_labels.size > 0) {
          const harnessLabelDir = await harnessDir.getDirectoryHandle('back_labels', { create: true });
          for (const [filename, blob] of hData.back_labels.entries()) {
            await writeBlob(harnessLabelDir, filename, blob);
          }
        }
        totalFiles += hData.images.size + hData.back_labels.size;
      }

      // 2. 項圈牽繩 (僅在有資料時輸出)
      const lData = filesToExport['項圈牽繩'];
      if (leashCount > 0 || leashExcelBuffer || lData.images.size > 0 || lData.back_labels.size > 0) {
        savedCategories.push('項圈牽繩');
        const leashDir = await dirHandle.getDirectoryHandle('項圈牽繩', { create: true });

        if (leashExcelBuffer) {
          const fh = await leashDir.getFileHandle('商品報價單_項圈 牽繩_auto_generate.xlsx', { create: true });
          const w = await fh.createWritable();
          await w.write(leashExcelBuffer);
          await w.close();
        }

        if (lData.images.size > 0) {
          const leashImgDir = await leashDir.getDirectoryHandle('images', { create: true });
          for (const [filename, blob] of lData.images.entries()) {
            await writeBlob(leashImgDir, filename, blob);
          }
        }
        if (lData.back_labels.size > 0) {
          const leashLabelDir = await leashDir.getDirectoryHandle('back_labels', { create: true });
          for (const [filename, blob] of lData.back_labels.entries()) {
            await writeBlob(leashLabelDir, filename, blob);
          }
        }
        totalFiles += lData.images.size + lData.back_labels.size;
      }

      if (savedCategories.length === 0) {
        alert('無可儲存的產品資料。');
        return;
      }

      setProgress(100, `已成功將【${savedCategories.join('、')}】之 ${totalFiles} 個圖檔與 Excel 存入所選資料夾！`);
      logMessage(`全部處理完成！已儲存【${savedCategories.join('、')}】資料，共 ${totalFiles} 個不重複圖檔（主圖/情境圖/尺寸規格表已自動去重複共用），無安全性警告！`, 'success');
    } catch (err) {
      console.error('寫入資料夾失敗:', err);
      logMessage(`寫入資料夾失敗: ${err.message}`, 'error');
    }
  });

  async function writeBlob(dirHandle, filename, blob) {
    try {
      const fileHandle = await dirHandle.getFileHandle(filename, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
    } catch(e) {
      console.error(`寫入 ${filename} 失敗:`, e);
    }
  }

  btnDownloadZip.addEventListener('click', async () => {
    if (processedResults.length === 0) return;

    progressContainer.classList.remove('hidden');
    setProgress(0, '正在打包為 ZIP 檔...');

    const zip = new JSZip();
    const exportedCategories = [];

    // 胸背帶 (僅在有資料時打包)
    const hData = filesToExport['胸背帶'];
    if (harnessCount > 0 || harnessExcelBuffer || hData.images.size > 0 || hData.back_labels.size > 0) {
      exportedCategories.push('胸背帶');
      const root = zip.folder('胸背帶');
      if (harnessExcelBuffer) {
        root.file('商品報價單_胸背帶_auto_generate.xlsx', harnessExcelBuffer);
      }
      if (hData.images.size > 0) {
        const imgFolder = root.folder('images');
        for (const [filename, blob] of hData.images.entries()) {
          imgFolder.file(filename, blob);
        }
      }
      if (hData.back_labels.size > 0) {
        const labelFolder = root.folder('back_labels');
        for (const [filename, blob] of hData.back_labels.entries()) {
          labelFolder.file(filename, blob);
        }
      }
    }

    // 項圈牽繩 (僅在有資料時打包)
    const lData = filesToExport['項圈牽繩'];
    if (leashCount > 0 || leashExcelBuffer || lData.images.size > 0 || lData.back_labels.size > 0) {
      exportedCategories.push('項圈牽繩');
      const root = zip.folder('項圈牽繩');
      if (leashExcelBuffer) {
        root.file('商品報價單_項圈 牽繩_auto_generate.xlsx', leashExcelBuffer);
      }
      if (lData.images.size > 0) {
        const imgFolder = root.folder('images');
        for (const [filename, blob] of lData.images.entries()) {
          imgFolder.file(filename, blob);
        }
      }
      if (lData.back_labels.size > 0) {
        const labelFolder = root.folder('back_labels');
        for (const [filename, blob] of lData.back_labels.entries()) {
          labelFolder.file(filename, blob);
        }
      }
    }

    if (exportedCategories.length === 0) {
      alert('無可供匯出的資料！');
      return;
    }

    setProgress(50, '正在壓縮...');
    const blob = await zip.generateAsync({ type: "blob" });
    setProgress(100, 'ZIP 打包完成！正在下載...');
    const filenameSuffix = exportedCategories.join('_');
    saveAs(blob, `MyFamily_Output_${filenameSuffix}_${new Date().toISOString().slice(0, 10)}.zip`);
    logMessage(`ZIP 檔案已下載（包含：${exportedCategories.join('、')}）`, 'success');
  });
});
