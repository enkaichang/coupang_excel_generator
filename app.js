const APP_VERSION = 'v1.4.6';

function normalizeHeaderKey(str) {
  if (!str) return '';
  return String(str)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[\uFEFF\xA0]/g, '')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .trim();
}

/**
 * Post-process xlsx buffer to inject missing <dimension> tags.
 * xlsx-populate strips <dimension> on output; some systems (e.g. Coupang) need it.
 */
async function patchXlsxDimension(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const sheetFiles = Object.keys(zip.files).filter(f => /^xl\/worksheets\/sheet\d+\.xml$/.test(f));

  for (const sheetFile of sheetFiles) {
    let xml = await zip.file(sheetFile).async('string');

    // Skip if <dimension> already exists
    if (/<dimension\s/.test(xml)) continue;

    // Extract all cell references from <c r="XX"> in <sheetData>
    const sheetDataMatch = xml.match(/<sheetData[\s\S]*?<\/sheetData>/);
    if (!sheetDataMatch) continue;

    const cellRefs = [];
    const cellRegex = /<c\s+[^>]*r="([A-Z]+)(\d+)"/g;
    let m;
    while ((m = cellRegex.exec(sheetDataMatch[0])) !== null) {
      cellRefs.push({ col: m[1], row: parseInt(m[2], 10) });
    }
    if (cellRefs.length === 0) continue;

    // Convert column letters to numbers for comparison
    const colToNum = (col) => {
      let n = 0;
      for (let i = 0; i < col.length; i++) {
        n = n * 26 + (col.charCodeAt(i) - 64);
      }
      return n;
    };
    const numToCol = (n) => {
      let s = '';
      while (n > 0) {
        n--;
        s = String.fromCharCode(65 + (n % 26)) + s;
        n = Math.floor(n / 26);
      }
      return s;
    };

    let minCol = Infinity, maxCol = 0, minRow = Infinity, maxRow = 0;
    for (const ref of cellRefs) {
      const cn = colToNum(ref.col);
      if (cn < minCol) minCol = cn;
      if (cn > maxCol) maxCol = cn;
      if (ref.row < minRow) minRow = ref.row;
      if (ref.row > maxRow) maxRow = ref.row;
    }

    const dimRef = `${numToCol(minCol)}${minRow}:${numToCol(maxCol)}${maxRow}`;
    const dimTag = `<dimension ref="${dimRef}"/>`;

    // Insert <dimension> right before <sheetViews>
    if (xml.includes('<sheetViews')) {
      xml = xml.replace('<sheetViews', dimTag + '<sheetViews');
    } else if (xml.includes('<sheetFormatPr')) {
      xml = xml.replace('<sheetFormatPr', dimTag + '<sheetFormatPr');
    }

    zip.file(sheetFile, xml);
  }

  return await zip.generateAsync({ 
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
}

function buildHeaderMap(worksheet, headerRow = 5) {
  const map = {};
  if (!worksheet) return map;
  const maxCol = worksheet.usedRange() ? worksheet.usedRange().endCell().columnNumber() : 70;
  for (let c = 1; c <= maxCol; c++) {
    const val = worksheet.cell(headerRow, c).value();
    if (val) {
      const raw = val.toString();
      const trimmed = raw.trim();
      const normLf = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const normLfTrim = normLf.trim();
      const normKey = normalizeHeaderKey(raw);
      const noSpace = raw.replace(/\s+/g, '');
      const cleanAll = normKey.replace(/\s+/g, '');

      if (!(raw in map)) map[raw] = c;
      if (!(trimmed in map)) map[trimmed] = c;
      if (!(normLf in map)) map[normLf] = c;
      if (!(normLfTrim in map)) map[normLfTrim] = c;
      if (!(normKey in map)) map[normKey] = c;
      if (!(noSpace in map)) map[noSpace] = c;
      if (!(cleanAll in map)) map[cleanAll] = c;
    }
  }
  return map;
}

function findHeaderColIdx(map, colName) {
  if (!colName || !map) return undefined;
  if (map[colName] !== undefined) return map[colName];
  if (map[colName.trim()] !== undefined) return map[colName.trim()];
  const normLf = colName.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (map[normLf] !== undefined) return map[normLf];
  if (map[normLf.trim()] !== undefined) return map[normLf.trim()];
  const normKey = normalizeHeaderKey(colName);
  if (map[normKey] !== undefined) return map[normKey];
  const noSpace = colName.replace(/\s+/g, '');
  if (map[noSpace] !== undefined) return map[noSpace];
  const cleanAll = normKey.replace(/\s+/g, '');
  if (map[cleanAll] !== undefined) return map[cleanAll];

  for (const [k, idx] of Object.entries(map)) {
    if (normalizeHeaderKey(k) === normKey || k.replace(/\s+/g, '') === noSpace || normalizeHeaderKey(k).replace(/\s+/g, '') === cleanAll) {
      return idx;
    }
  }
  return undefined;
}

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
  
  const inputRowStart = document.getElementById('inputRowStart');
  const inputRowEnd = document.getElementById('inputRowEnd');
  const btnResetRange = document.getElementById('btnResetRange');
  const rangeInfoHint = document.getElementById('rangeInfoHint');

  if (inputRowStart) {
    inputRowStart.addEventListener('input', () => updateRangeHintUI());
  }
  if (inputRowEnd) {
    inputRowEnd.addEventListener('input', () => updateRangeHintUI());
  }
  if (btnResetRange) {
    btnResetRange.addEventListener('click', () => {
      if (inputRowStart) inputRowStart.value = currentConfig.source.row_start || 4;
      if (inputRowEnd) inputRowEnd.value = '';
      updateRangeHintUI();
    });
  }
  
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
  const guiCategoryRuleForm = document.getElementById('guiCategoryRuleForm');
  const btnAddCategoryRule = document.getElementById('btnAddCategoryRule');
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
        <div class="input-row"><label>篩選欄位(有填寫才處理):</label><input type="text" id="cfg_filter_column" value="${config.source.filter_column || ''}"></div>
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
        file_path: "商品資料.xlsx",
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

  function renderGuiCategoryRules(rules) {
    if (!guiCategoryRuleForm) return;
    guiCategoryRuleForm.innerHTML = '';
    for (const rule of rules) {
      addCategoryRuleRow(rule);
    }
  }

  function addCategoryRuleRow(rule = {}) {
    if (!guiCategoryRuleForm) return;
    const name = rule.name || '';
    const keywords = Array.isArray(rule.keywords) ? rule.keywords.join(', ') : (rule.keywords || '');
    const template_type = (rule.template_type || 'LEASH').toUpperCase();
    const category_name = rule.category_name || '';
    const subfolder = rule.subfolder || '';

    const div = document.createElement('div');
    div.className = 'category-rule-card';
    div.innerHTML = `
      <div class="category-rule-header">
        <div class="category-rule-title">
          <span>📦 品類名稱:</span>
          <input type="text" class="rule-name-input" placeholder="規則名稱 (例: 胸背帶)" value="${name}" style="padding:4px 10px; font-size:0.88rem; font-weight:600; width:150px; border:1px solid #cbd5e1; border-radius:6px;">
        </div>
        <button class="btn btn-danger btn-sm" onclick="this.closest('.category-rule-card').remove()">刪除規則</button>
      </div>
      <div class="category-rule-grid">
        <div class="rule-field-group">
          <label>匹配關鍵字 (TYPE 欄位 / 中文品名關鍵字，逗號分隔):</label>
          <input type="text" class="rule-keywords-input" placeholder="例: HARNESS, 胸背帶, 背帶" value="${keywords}">
        </div>
        <div class="rule-field-group">
          <label>套用 Excel 模板種類:</label>
          <select class="rule-template-select">
            <option value="HARNESS" ${template_type === 'HARNESS' ? 'selected' : ''}>胸背帶模板 (商品報價單_胸背帶.xlsx)</option>
            <option value="LEASH" ${template_type === 'LEASH' ? 'selected' : ''}>項圈牽繩模板 (商品報價單_項圈 牽繩.xlsx)</option>
          </select>
        </div>
        <div class="rule-field-group">
          <label>酷澎目標細分商品種類 (分類代碼及名稱):</label>
          <input type="text" class="rule-category-input" placeholder="例: 寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)" value="${category_name}">
        </div>
        <div class="rule-field-group">
          <label>輸出子資料夾名稱:</label>
          <input type="text" class="rule-subfolder-input" placeholder="例: 胸背帶 或 項圈牽繩" value="${subfolder}">
        </div>
      </div>
    `;
    guiCategoryRuleForm.appendChild(div);
  }

  if (btnAddCategoryRule) {
    btnAddCategoryRule.addEventListener('click', () => {
      addCategoryRuleRow({
        name: '新品類',
        keywords: '',
        template_type: 'LEASH',
        category_name: '',
        subfolder: '新品類'
      });
    });
  }

  function parseGuiCategoryRules() {
    if (!guiCategoryRuleForm) return currentCategoryRules;
    const cards = guiCategoryRuleForm.querySelectorAll('.category-rule-card');
    const rules = [];
    cards.forEach(card => {
      const name = card.querySelector('.rule-name-input').value.trim();
      const keywordsStr = card.querySelector('.rule-keywords-input').value;
      const template_type = card.querySelector('.rule-template-select').value;
      const category_name = card.querySelector('.rule-category-input').value.trim();
      const subfolder = card.querySelector('.rule-subfolder-input').value.trim();
      
      const keywords = keywordsStr.split(',').map(s => s.trim()).filter(Boolean);
      if (name || keywords.length > 0) {
        rules.push({
          name: name || '未命名品類',
          keywords: keywords,
          template_type: template_type,
          template_name: template_type === 'HARNESS' ? '商品報價單_胸背帶.xlsx' : '商品報價單_項圈 牽繩.xlsx',
          category_name: category_name,
          subfolder: subfolder || name || '未分類'
        });
      }
    });
    return rules.length > 0 ? rules : window.AppConfig.getDefaultCategoryRules();
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
  let currentCategoryRules = window.AppConfig.getCategoryRules();
  let currentCollectionAliases = window.AppConfig.getCollectionAliases();
  let currentColorAliases = window.AppConfig.getColorAliases();

  // Load custom templates if available
  let customTemplates = { HARNESS: null, LEASH: null };
  try {
    const savedTemplates = localStorage.getItem('coupang_templates') || localStorage.getItem('my_family_templates');
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
      excelData = window.MyFamilyProcessor.extractMappingsFromExcel(loadedWorkbook, headerRow, rowStart, filterColName, currentConfig?.source?.sheet_name);
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

  function getSourceSheet(wb) {
    if (!wb) return null;
    const headerRow = currentConfig?.source?.header_row || 3;
    const targetSheetName = currentConfig?.source?.sheet_name;

    // 1. 若設定有指定目標工作表名稱，且工作表存在且有資料，直接採用
    if (targetSheetName) {
      const s = wb.sheet(targetSheetName);
      if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
        return s;
      }
    }

    // 2. 搜尋常見工作表名稱，優先選擇有實質資料列 (> headerRow) 的工作表
    const candidateNames = ['商品資料', 'MYFAMILY', 'MY FAMILY', 'My Family', '工作表1', 'Sheet1'];
    for (const name of candidateNames) {
      const s = wb.sheet(name);
      if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
        return s;
      }
    }

    // 3. 搜尋所有工作表中列數大於表頭且資料最多的工作表
    const allSheets = wb.sheets ? wb.sheets() : [];
    let bestSheet = null;
    let maxRows = 0;
    for (const s of allSheets) {
      if (s && s.usedRange()) {
        const rows = s.usedRange().endCell().rowNumber();
        if (rows > headerRow && rows > maxRows) {
          maxRows = rows;
          bestSheet = s;
        }
      }
    }
    if (bestSheet) return bestSheet;

    // 4. 退回指定工作表或候選名稱或第 0 個工作表
    if (targetSheetName) {
      const s = wb.sheet(targetSheetName);
      if (s) return s;
    }
    for (const name of candidateNames) {
      const s = wb.sheet(name);
      if (s) return s;
    }
    return wb.sheet(0);
  }

  function getValidRowsInfo(ws, customStart = null, customEnd = null) {
    if (!ws || !ws.usedRange()) {
      return {
        validRowIndices: [],
        count: 0,
        totalRows: 0,
        headerRow: 3,
        start: 4,
        end: 0,
        headers: {},
        filterColIdx: undefined,
        nameColIdx: undefined,
        filterColName: currentConfig?.source?.filter_column || '中文背標'
      };
    }

    const headerRow = currentConfig?.source?.header_row || 3;
    const totalRows = ws.usedRange().endCell().rowNumber();

    let start = (customStart !== null && customStart !== undefined && customStart !== '')
      ? parseInt(customStart)
      : (parseInt(inputRowStart ? inputRowStart.value : '') || currentConfig?.source?.row_start || 4);
    if (isNaN(start) || start < 1) start = 4;

    let end = (customEnd !== null && customEnd !== undefined && customEnd !== '')
      ? parseInt(customEnd)
      : (inputRowEnd && inputRowEnd.value.trim() ? parseInt(inputRowEnd.value.trim()) : totalRows);
    if (isNaN(end) || end < 1) end = totalRows;
    end = Math.min(end, totalRows);

    const headers = buildHeaderMap(ws, headerRow);

    const filterColName = currentConfig?.source?.filter_column || '中文背標';
    const filterColIdx = findHeaderColIdx(headers, filterColName);
    const nameColIdx = findHeaderColIdx(headers, '中文品名') || findHeaderColIdx(headers, '商品名稱') || findHeaderColIdx(headers, '品名');

    const validRowIndices = [];
    if (start <= end) {
      for (let r = start; r <= end; r++) {
        if (filterColIdx) {
          const filterVal = (getCellValue(ws.cell(r, filterColIdx)) || '').toString().trim();
          if (!filterVal) continue;
        }
        const firstCell = ws.cell(r, 1);
        const fill = firstCell.style('fill');
        if (fill && fill.type === 'solid' && fill.color && typeof fill.color === 'string') {
          if (fill.color.toUpperCase().includes('FAD9D6')) continue;
        }
        const zhName = nameColIdx ? (getCellValue(ws.cell(r, nameColIdx)) || '').toString().trim() : '';
        if (!zhName) continue;
        validRowIndices.push(r);
      }
    }

    return {
      validRowIndices,
      count: validRowIndices.length,
      totalRows,
      headerRow,
      start,
      end,
      headers,
      filterColIdx,
      nameColIdx,
      filterColName
    };
  }

  function calculateValidProducts(ws, customStart = null, customEnd = null) {
    return getValidRowsInfo(ws, customStart, customEnd);
  }

  function updateRangeHintUI() {
    if (!loadedWorkbook) return;
    const ws = getSourceSheet(loadedWorkbook);
    if (!ws || !ws.usedRange()) return;

    const stats = calculateValidProducts(ws);
    if (rangeInfoHint) {
      rangeInfoHint.textContent = `來源工作表「${ws.name()}」共 ${stats.totalRows} 列（有效商品：${stats.count} 筆，表頭在第 ${stats.headerRow} 列）`;
      rangeInfoHint.classList.add('active');
    }
    if (excelFileInfo && sourceExcelFile) {
      excelFileInfo.innerHTML = `已選擇: <strong>${sourceExcelFile.name}</strong> <span style="margin-left: 8px; padding: 3px 12px; background: rgba(99, 102, 241, 0.15); color: #4f46e5; border: 1px solid rgba(99, 102, 241, 0.35); border-radius: 9999px; font-weight: 700; font-size: 0.85rem;">有效商品：${stats.count} 筆</span>`;
      excelFileInfo.classList.remove('hidden');
    }
  }

  async function handleExcelFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      alert('請上傳 Excel 檔案 (.xlsx 或 .xls)');
      return;
    }
    sourceExcelFile = file;
    excelDropZone.classList.add('has-file');
    excelFileInfo.textContent = `已選擇: ${file.name}`;
    excelFileInfo.classList.remove('hidden');
    checkReady();

    try {
      const arrayBuffer = await file.arrayBuffer();
      loadedWorkbook = await XlsxPopulate.fromDataAsync(arrayBuffer);
      
      const ws = getSourceSheet(loadedWorkbook);
      if (ws && ws.usedRange()) {
        const totalR = ws.usedRange().endCell().rowNumber();
        if (inputRowEnd) {
          inputRowEnd.placeholder = `最大 ${totalR} 列`;
        }
        updateRangeHintUI();
        const stats = calculateValidProducts(ws);
        logMessage(`已載入來源 Excel: ${file.name}（工作表「${ws.name()}」共 ${totalR} 列，符合條件之有效商品共 ${stats.count} 筆）`, 'success');
      } else {
        logMessage(`已載入來源 Excel: ${file.name}`);
      }

      await autoScanAndSyncMappings(true);
    } catch (err) {
      console.warn('解析 Excel 欄位對照失敗:', err);
      logMessage(`解析 Excel 失敗: ${err.message}`, 'error');
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
    renderGuiCategoryRules(currentCategoryRules);
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
      currentCategoryRules = parseGuiCategoryRules();
      currentCollectionAliases = parseGuiCollection();
      currentColorAliases = parseGuiColor();
      
      localStorage.setItem('coupang_config', JSON.stringify(currentConfig));
      localStorage.setItem('coupang_category_rules', JSON.stringify(currentCategoryRules));
      localStorage.setItem('coupang_collection_aliases', JSON.stringify(currentCollectionAliases));
      localStorage.setItem('coupang_color_aliases', JSON.stringify(currentColorAliases));
      localStorage.setItem('my_family_config', JSON.stringify(currentConfig));
      localStorage.setItem('my_family_category_rules', JSON.stringify(currentCategoryRules));
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
        localStorage.setItem('coupang_templates', JSON.stringify(customTemplates));
        localStorage.setItem('my_family_templates', JSON.stringify(customTemplates));
      }

      updateRangeHintUI();
      alert('設定已儲存！');
      configModal.classList.add('hidden');
    } catch (e) {
      alert('JSON 格式錯誤，請檢查後再儲存。\n錯誤訊息: ' + e.message);
    }
  });

  btnResetConfig.addEventListener('click', () => {
    if (confirm('確定要還原為系統預設值嗎？')) {
      localStorage.removeItem('coupang_config');
      localStorage.removeItem('coupang_category_rules');
      localStorage.removeItem('coupang_collection_aliases');
      localStorage.removeItem('coupang_color_aliases');
      localStorage.removeItem('coupang_templates');
      localStorage.removeItem('my_family_config');
      localStorage.removeItem('my_family_category_rules');
      localStorage.removeItem('my_family_collection_aliases');
      localStorage.removeItem('my_family_color_aliases');
      localStorage.removeItem('my_family_templates');
      
      currentConfig = window.AppConfig.getDefaultConfig();
      currentCategoryRules = window.AppConfig.getDefaultCategoryRules();
      currentCollectionAliases = window.AppConfig.getDefaultCollectionAliases();
      currentColorAliases = window.AppConfig.getDefaultColorAliases();
      customTemplates = { HARNESS: null, LEASH: null };
      
      renderGuiConfig(currentConfig);
      renderGuiCategoryRules(currentCategoryRules);
      renderGuiCollection(currentCollectionAliases);
      renderGuiColor(currentColorAliases);
      
      templateHarnessStatus.textContent = "已套用預設模板";
      templateLeashStatus.textContent = "已套用預設模板";
      templateHarnessInput.value = "";
      templateLeashInput.value = "";
      if (inputRowStart) inputRowStart.value = currentConfig.source.row_start || 4;
      if (inputRowEnd) inputRowEnd.value = "";
      
      updateRangeHintUI();
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
    saveAs(content, "Coupang_Sample_Folders.zip");
  });

  let filesToExport = {};

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
    filesToExport = {};
    
    statTotalSku.textContent = '0';
    statMatchedImages.textContent = '0';
    statMissingImages.textContent = '0';
    
    try {
      logMessage('開始解析來源 Excel...');
      const arrayBuffer = await sourceExcelFile.arrayBuffer();
      const workbook = loadedWorkbook || await XlsxPopulate.fromDataAsync(arrayBuffer);
      
      const ws = getSourceSheet(workbook);
      if (!ws) throw new Error('找不到任何工作表，請確認 Excel 檔案格式');
      logMessage(`使用來源工作表: 「${ws.name()}」`);

      if (!ws.usedRange()) throw new Error('來源工作表似乎是空的');

      const validInfo = getValidRowsInfo(ws);
      const {
        validRowIndices,
        count: totalValidCount,
        totalRows,
        headerRow,
        start: rowStart,
        end: rowEnd,
        headers,
        filterColIdx,
        nameColIdx,
        filterColName
      } = validInfo;

      if (rowStart > totalRows) {
        throw new Error(`設定的起始列號 (${rowStart}) 超過來源工作表的最大列數 (${totalRows})`);
      }
      if (rowStart > rowEnd) {
        throw new Error(`設定的起始列號 (${rowStart}) 不能大於結束列號 (${rowEnd})`);
      }

      logMessage(`已讀取表頭欄位: 共 ${Object.keys(headers).length} 欄 (第 ${headerRow} 列)`);
      if (filterColIdx) {
        logMessage(`篩選設定: 依「${filterColName}」(第 ${filterColIdx} 欄) 篩選非空值商品`);
      } else {
        logMessage(`篩選設定: 表頭未找到「${filterColName}」欄位，不進行該欄位篩選`, 'info');
      }

      logMessage(`處理範圍：第 ${rowStart} 列 至 第 ${rowEnd} 列（共掃描 ${rowEnd - rowStart + 1} 列，其中符合條件之有效商品共 ${totalValidCount} 筆，工作表總列數: ${totalRows} 列）`);

      if (totalValidCount === 0) {
        setProgress(100, '所選範圍內無符合條件之有效商品資料。');
        logMessage('未找到任何符合篩選條件的有效商品資料。', 'warning');
        btnSaveToFolder.disabled = true;
        btnDownloadZip.disabled = true;
        return;
      }

      const processor = new window.MyFamilyProcessor(
        photoFilesArray,
        arrayBuffer,
        currentColorAliases,
        window.AppConfig.zhColorMap,
        window.AppConfig.targetCombos,
        currentCollectionAliases,
        currentCategoryRules
      );

      // Load Templates (Custom or Built-in) using XlsxPopulate to preserve EXACT formatting (dropdowns, validations)
      const harnessB64 = customTemplates.HARNESS || window.CoupangTemplates?.HARNESS || window.MyFamilyTemplates?.HARNESS;
      const leashB64 = customTemplates.LEASH || window.CoupangTemplates?.LEASH || window.MyFamilyTemplates?.LEASH;
      
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
      
      const harnessMap = buildHeaderMap(outWsHarness, 5);
      const leashMap = buildHeaderMap(outWsLeash, 5);
      
      let harnessRowIdx = 9;
      let leashRowIdx = 9;
      
      let totalCount = 0;
      let matchCount = 0;
      let missCount = 0;
      
      const imageCounters = {};
      const prevBaseNames = {};
      const lastMainImgNames = {};
      const lastSc1ImgNames = {};
      const lastSc2ImgNames = {};
      const lastChartImgNames = {};
      
      for (let i = 0; i < validRowIndices.length; i++) {
        const r = validRowIndices[i];
        const currentItemNum = i + 1;

        const zhName = nameColIdx ? (getCellValue(ws.cell(r, nameColIdx)) || '').toString().trim() : '';
        if (!zhName) continue;

        const skuIdx = findHeaderColIdx(headers, 'SKU') || findHeaderColIdx(headers, '商品條碼') || findHeaderColIdx(headers, '條碼') || findHeaderColIdx(headers, 'EAN');
        const skuVal = skuIdx ? getCellValue(ws.cell(r, skuIdx)) : '';
        const skuStr = formatBarcode(skuVal);
        const collIdx = findHeaderColIdx(headers, 'COLLECTION') || findHeaderColIdx(headers, '系列') || findHeaderColIdx(headers, 'Collection');
        const rawCollection = collIdx ? (getCellValue(ws.cell(r, collIdx)) || '').toString().trim() : '';
        const typeIdx = findHeaderColIdx(headers, 'TYPE') || findHeaderColIdx(headers, '種類') || findHeaderColIdx(headers, '品類') || findHeaderColIdx(headers, 'Type');
        const rawType = typeIdx ? (getCellValue(ws.cell(r, typeIdx)) || '').toString().trim() : '';
        const colorIdx = findHeaderColIdx(headers, 'COLOR') || findHeaderColIdx(headers, '顏色') || findHeaderColIdx(headers, 'Color');
        const rawColor = colorIdx ? (getCellValue(ws.cell(r, colorIdx)) || '').toString().trim() : '';
        const sizeIdx = findHeaderColIdx(headers, 'SIZE') || findHeaderColIdx(headers, '尺寸') || findHeaderColIdx(headers, 'Size');
        const rawSize = sizeIdx ? (getCellValue(ws.cell(r, sizeIdx)) || '').toString().trim() : '';
        
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

        const targetSubFolder = targetInfo.target_subfolder || '項圈牽繩';
        const tmplType = (targetInfo.template_type || 'LEASH').toUpperCase();

        if (!filesToExport[targetSubFolder]) {
          filesToExport[targetSubFolder] = {
            images: new Map(),
            back_labels: new Map(),
            hasHarnessExcel: false,
            hasLeashExcel: false
          };
        }
        if (!imageCounters[targetSubFolder]) imageCounters[targetSubFolder] = 1;
        if (prevBaseNames[targetSubFolder] === undefined) prevBaseNames[targetSubFolder] = '';
        if (lastMainImgNames[targetSubFolder] === undefined) lastMainImgNames[targetSubFolder] = '';
        if (lastSc1ImgNames[targetSubFolder] === undefined) lastSc1ImgNames[targetSubFolder] = '';
        if (lastSc2ImgNames[targetSubFolder] === undefined) lastSc2ImgNames[targetSubFolder] = '';
        if (lastChartImgNames[targetSubFolder] === undefined) lastChartImgNames[targetSubFolder] = '';

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
              const blob = await window.ImageUtils.resizeAndPad(await window.ImageUtils.loadImage(imgs.main), 1000, 1000, '#FFFFFF');
              mainImgName = `${cleanName}主圖.jpg`;
              filesToExport[targetSubFolder].images.set(mainImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 主圖處理失敗: ${imgs.main.name}`, 'warning');
            }
          }
          lastMainImgNames[targetSubFolder] = mainImgName;

          if (imgs.sc1) {
            try {
              const blob = await window.ImageUtils.resizeAndPad(await window.ImageUtils.loadImage(imgs.sc1), 1000, 1000, '#FFFFFF');
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
              const blob = await window.ImageUtils.resizeAndPad(await window.ImageUtils.loadImage(imgs.sc2), 1000, 1000, '#FFFFFF');
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
            const szSuffix = (rawSize || parsed.size) ? `_${String(rawSize || parsed.size).trim()}` : '';
            const cleanLabelName = sanitizeFilename(`背標_${zhName}${szSuffix}`);
            labelImgName = `${cleanLabelName}.jpg`;
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
        if (tmplType === 'HARNESS') {
          targetWs = outWsHarness;
          targetRowIdx = harnessRowIdx++;
          tmplMap = harnessMap;
          harnessCount++;
          filesToExport[targetSubFolder].hasHarnessExcel = true;
        } else {
          targetWs = outWsLeash;
          targetRowIdx = leashRowIdx++;
          tmplMap = leashMap;
          leashCount++;
          filesToExport[targetSubFolder].hasLeashExcel = true;
        }
        
        const getSourceHeaderIdx = (colName) => {
          return findHeaderColIdx(headers, colName);
        };

        const setVal = (colName, val) => {
          const colIdx = findHeaderColIdx(tmplMap, colName);
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

        for (const [tKey, sCol] of Object.entries(currentConfig.field_mappings.dynamic || {})) {
            if (['細分商品種類', '顏色', '尺寸'].includes(tKey)) continue;
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

        // 品牌、系列、顏色、尺寸一律使用中文品名解析之結果（確保匯出為中文品名所定義之中文顏色與尺寸）
        const finalColor = parsed.color || processor.extractColorFromZhName(zhName, rawSize);
        const finalSize = parsed.size || rawSize;
        if (parsed.brand) {
          setVal('品牌', parsed.brand);
          if (!currentConfig.field_mappings.fixed?.['製造廠商']) {
            setVal('製造廠商', parsed.brand);
          }
        }
        if (parsed.collection) setVal('系列', parsed.collection);
        if (finalColor) setVal('顏色', finalColor);
        if (finalSize) setVal('尺寸', finalSize);
        setVal('商品名稱', zhName);

        if (mainImgName) setVal('商品正面(主要圖片）', mainImgName);
        if (sc1ImgName) setVal('商品側拍或情境圖 1\n(消費者可見圖片）', sc1ImgName);
        if (sc2ImgName) setVal('商品側拍或情境圖 2\n(消費者可見圖片）', sc2ImgName);
        if (chartImgName) {
          const chartColIdx = findHeaderColIdx(tmplMap, '商品詳細說明圖集\n(消費者可見圖片）') ||
                              findHeaderColIdx(tmplMap, '詳細說明圖集') ||
                              findHeaderColIdx(tmplMap, '尺寸圖') ||
                              findHeaderColIdx(tmplMap, '尺寸規格');
          if (chartColIdx) {
            targetWs.cell(targetRowIdx, chartColIdx).value(chartImgName);
          }
        }
        if (labelImgName) {
          const labelColIdx = findHeaderColIdx(tmplMap, '商品實際中文背標圖\n(內部上架審核用)') ||
                              findHeaderColIdx(tmplMap, '中文背標圖') ||
                              findHeaderColIdx(tmplMap, '背標');
          if (labelColIdx) {
            targetWs.cell(targetRowIdx, labelColIdx).value(labelImgName);
          }
        }

        processedResults.push({
          skuStr, zhName, subFolder: targetSubFolder
        });

        if (currentItemNum % 2 === 0 || currentItemNum === totalValidCount) {
          const pct = Math.min(Math.round((currentItemNum / totalValidCount) * 100), 100);
          setProgress(pct, `處理商品中 (${currentItemNum} / ${totalValidCount})...`);
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
        harnessExcelBuffer = await patchXlsxDimension(harnessExcelBuffer);
      } else {
        harnessExcelBuffer = null;
      }

      if (leashCount > 0) {
        leashExcelBuffer = await outWbLeash.outputAsync();
        leashExcelBuffer = await patchXlsxDimension(leashExcelBuffer);
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
        if (harnessCount > 0) typesDesc.push(`胸背帶模板: ${harnessCount} 筆`);
        if (leashCount > 0) typesDesc.push(`項圈牽繩模板: ${leashCount} 筆`);
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
    setProgress(0, '正在計算寫入檔案總數...');

    try {
      const savedCategories = [];
      let totalFiles = 0;
      let totalWriteCount = 0;

      // 預先計算總檔案數（Excel 檔案 + 圖片 + 背標）
      for (const [folderName, folderData] of Object.entries(filesToExport)) {
        const hasContent = folderData.images.size > 0 || folderData.back_labels.size > 0 || folderData.hasHarnessExcel || folderData.hasLeashExcel;
        if (!hasContent) continue;
        if (folderData.hasHarnessExcel && harnessExcelBuffer) totalWriteCount++;
        if (folderData.hasLeashExcel && leashExcelBuffer) totalWriteCount++;
        totalWriteCount += folderData.images.size + folderData.back_labels.size;
      }

      if (totalWriteCount === 0) {
        alert('無可儲存的產品資料。');
        return;
      }

      let writtenCount = 0;
      const updateWriteProgress = async (fileLabel) => {
        writtenCount++;
        const pct = Math.min(Math.round((writtenCount / totalWriteCount) * 100), 100);
        setProgress(pct, `正在寫入檔案 (${writtenCount} / ${totalWriteCount}): ${fileLabel}...`);
        if (writtenCount % 3 === 0 || writtenCount === totalWriteCount) {
          await new Promise(r => setTimeout(r, 0));
        }
      };

      for (const [folderName, folderData] of Object.entries(filesToExport)) {
        const hasContent = folderData.images.size > 0 || folderData.back_labels.size > 0 || folderData.hasHarnessExcel || folderData.hasLeashExcel;
        if (!hasContent) continue;

        savedCategories.push(folderName);
        const subDir = await dirHandle.getDirectoryHandle(folderName, { create: true });
        
        if (folderData.hasHarnessExcel && harnessExcelBuffer) {
          const fname = '商品報價單_胸背帶_auto_generate.xlsx';
          const fh = await subDir.getFileHandle(fname, { create: true });
          const w = await fh.createWritable();
          await w.write(harnessExcelBuffer);
          await w.close();
          await updateWriteProgress(`${folderName}/${fname}`);
        }

        if (folderData.hasLeashExcel && leashExcelBuffer) {
          const fname = '商品報價單_項圈 牽繩_auto_generate.xlsx';
          const fh = await subDir.getFileHandle(fname, { create: true });
          const w = await fh.createWritable();
          await w.write(leashExcelBuffer);
          await w.close();
          await updateWriteProgress(`${folderName}/${fname}`);
        }

        if (folderData.images.size > 0) {
          const imgDir = await subDir.getDirectoryHandle('images', { create: true });
          for (const [filename, blob] of folderData.images.entries()) {
            await writeBlob(imgDir, filename, blob);
            await updateWriteProgress(`${folderName}/images/${filename}`);
          }
        }

        if (folderData.back_labels.size > 0) {
          const labelDir = await subDir.getDirectoryHandle('back_labels', { create: true });
          for (const [filename, blob] of folderData.back_labels.entries()) {
            await writeBlob(labelDir, filename, blob);
            await updateWriteProgress(`${folderName}/back_labels/${filename}`);
          }
        }

        totalFiles += folderData.images.size + folderData.back_labels.size;
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
      const safeName = sanitizeFilename(filename.replace(/\.[^/.]+$/, '')) + (filename.includes('.') ? filename.substring(filename.lastIndexOf('.')) : '');
      const fileHandle = await dirHandle.getFileHandle(safeName || filename, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
    } catch(e) {
      console.error(`寫入 ${filename} 失敗:`, e);
      logMessage(`[錯誤] 寫入 ${filename} 失敗: ${e.message || e}`, 'error');
    }
  }

  btnDownloadZip.addEventListener('click', async () => {
    if (processedResults.length === 0) return;

    progressContainer.classList.remove('hidden');
    setProgress(0, '正在準備 ZIP 檔案結構...');

    const zip = new JSZip();
    const exportedCategories = [];

    for (const [folderName, folderData] of Object.entries(filesToExport)) {
      const hasContent = folderData.images.size > 0 || folderData.back_labels.size > 0 || folderData.hasHarnessExcel || folderData.hasLeashExcel;
      if (!hasContent) continue;

      exportedCategories.push(folderName);
      const root = zip.folder(folderName);
      
      if (folderData.hasHarnessExcel && harnessExcelBuffer) {
        root.file('商品報價單_胸背帶_auto_generate.xlsx', harnessExcelBuffer);
      }
      if (folderData.hasLeashExcel && leashExcelBuffer) {
        root.file('商品報價單_項圈 牽繩_auto_generate.xlsx', leashExcelBuffer);
      }
      if (folderData.images.size > 0) {
        const imgFolder = root.folder('images');
        for (const [filename, blob] of folderData.images.entries()) {
          imgFolder.file(filename, blob);
        }
      }
      if (folderData.back_labels.size > 0) {
        const labelFolder = root.folder('back_labels');
        for (const [filename, blob] of folderData.back_labels.entries()) {
          labelFolder.file(filename, blob);
        }
      }
    }

    if (exportedCategories.length === 0) {
      alert('無可供匯出的資料！');
      return;
    }

    setProgress(5, '正在壓縮打包 ZIP 檔 (0%)...');
    const blob = await zip.generateAsync({ type: "blob" }, (metadata) => {
      const pct = Math.min(Math.max(Math.round(metadata.percent), 5), 99);
      const currentFile = metadata.currentFile ? ` (${metadata.currentFile})` : '';
      setProgress(pct, `正在壓縮打包 ZIP 檔 (${pct}%)${currentFile}...`);
    });
    setProgress(100, 'ZIP 打包完成！正在下載...');
    const filenameSuffix = exportedCategories.join('_');
    saveAs(blob, `Coupang_Output_${filenameSuffix}_${new Date().toISOString().slice(0, 10)}.zip`);
    logMessage(`ZIP 檔案已下載（包含：${exportedCategories.join('、')}）`, 'success');
  });
});
