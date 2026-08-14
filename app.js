const APP_VERSION = 'v1.8.4';

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

/**
 * Scan template excel headers (Row 5 = name, Row 6 = requirement) and match against baseline
 */
async function scanTemplateHeaders(arrayBuffer) {
  const wb = await XlsxPopulate.fromDataAsync(arrayBuffer);
  const ws = wb.sheets().find(s => s.name().startsWith('QF_')) || wb.sheet(0);
  const maxCol = ws.usedRange() ? ws.usedRange().endCell().columnNumber() : 70;
  const baseline = window.AppConfig.getBaselineMappings();
  const requiredColumns = [];

  for (let c = 1; c <= maxCol; c++) {
    const colNameVal = ws.cell(5, c).value();
    const reqVal = ws.cell(6, c).value();
    if (!colNameVal) continue;
    const colName = colNameVal.toString().trim();
    const reqStr = (reqVal || '').toString().trim();

    if (reqStr === '必填') {
      let status = 'new';
      let mappingType = 'dynamic';
      let defaultValue = '';

      // Check if system managed
      const isSystem = baseline.systemFields.some(sf => normalizeHeaderKey(sf) === normalizeHeaderKey(colName) || sf.replace(/\s+/g, '') === colName.replace(/\s+/g, ''));
      if (isSystem) {
        status = 'system';
        mappingType = 'system';
        defaultValue = '(系統自動處理)';
      } else {
        // Check dynamic baseline
        let foundDyn = null;
        for (const [k, v] of Object.entries(baseline.dynamic)) {
          if (normalizeHeaderKey(k) === normalizeHeaderKey(colName) || k.replace(/\s+/g, '') === colName.replace(/\s+/g, '')) {
            foundDyn = v;
            break;
          }
        }
        if (foundDyn !== null) {
          status = 'inherited';
          mappingType = 'dynamic';
          defaultValue = foundDyn;
        } else {
          // Check fixed baseline
          let foundFix = null;
          for (const [k, v] of Object.entries(baseline.fixed)) {
            if (normalizeHeaderKey(k) === normalizeHeaderKey(colName) || k.replace(/\s+/g, '') === colName.replace(/\s+/g, '')) {
              foundFix = v;
              break;
            }
          }
          if (foundFix !== null) {
            status = 'inherited';
            mappingType = 'fixed';
            defaultValue = foundFix;
          } else {
            status = 'new';
            mappingType = 'dynamic';
            defaultValue = '';
          }
        }
      }

      requiredColumns.push({
        colIdx: c,
        colName: colName,
        rawName: colNameVal.toString(),
        status: status,
        mappingType: mappingType,
        defaultValue: defaultValue
      });
    }
  }

  return {
    sheetName: ws.name(),
    requiredColumns
  };
}

function b64ToArrayBuffer(base64) {
  const binary_string = window.atob(base64);
  const len = binary_string.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary_string.charCodeAt(i);
  }
  return bytes.buffer;
}

function arrayBufferToB64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}

function readAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = error => reject(error);
    reader.readAsDataURL(file);
  });
}

function sanitizeFilename(name) {
  if (!name) return '';
  return String(name).replace(/[\\/:*?"<>|]/g, '_').trim();
}

function formatBarcode(val) {
  if (val === null || val === undefined || val === '') return '';
  let s = String(val).trim();
  if (/^[0-9]+\.0+$/.test(s)) {
    s = s.split('.')[0];
  }
  if (/^[0-9]+(\.[0-9]+)?e\+[0-9]+$/i.test(s)) {
    try {
      const num = Number(s);
      if (!isNaN(num)) {
        s = BigInt(Math.round(num)).toString();
      }
    } catch(e) {}
  }
  return s;
}
function getCellValue(cell) {
  if (!cell) return '';
  const val = cell.value();
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && val.text) return val.text;
  return val;
}

function getBaseProductName(fullName, size) {
  if (!fullName) return '';
  let s = String(fullName).trim();
  if (size) {
    const szStr = String(size).trim();
    if (szStr && s.endsWith(szStr)) {
      s = s.slice(0, -szStr.length).trim();
    }
  }
  return s.replace(/\s+([0-9]*X*[SML]+|[0-9]+(?:cm)?\/[0-9.]+(?:mm|cm)?|[0-9]+-[0-9.]+|[0-9]+\/[0-9]+mm)$/i, '').trim();
}

/**
 * Scan data rows in source Excel to dynamically detect which template profiles are actively needed
 */
function detectRequiredTemplates(ws, headerRow, rowStart, maxRow, profiles, sourceConfig = null) {
  if (!ws || !profiles || profiles.length === 0) return profiles || [];
  const headerMap = buildHeaderMap(ws, headerRow);
  let nameColIdx = findHeaderColIdx(headerMap, '商品名稱') || findHeaderColIdx(headerMap, '中文品名') || findHeaderColIdx(headerMap, '品名') || findHeaderColIdx(headerMap, 'NAME') || findHeaderColIdx(headerMap, '中文') || findHeaderColIdx(headerMap, '商品');
  const typeColIdx = findHeaderColIdx(headerMap, 'TYPE') || findHeaderColIdx(headerMap, '種類') || findHeaderColIdx(headerMap, '品類') || findHeaderColIdx(headerMap, 'Type');
  const filterColName = sourceConfig?.filter_column || '中文背標';
  const filterColIdx = findHeaderColIdx(headerMap, filterColName);

  const matchedProfileIds = new Set();
  const processor = new window.MyFamilyProcessor([], null, {}, {}, profiles);

  const range = ws.usedRange();
  const totalR = range ? range.endCell().rowNumber() : 100;
  const actualStart = Math.max(headerRow + 1, rowStart || (headerRow + 1));
  const actualEnd = maxRow ? Math.min(maxRow, totalR) : totalR;

  // Fallback: discover name column by scanning cell contents against known category keywords
  if (!nameColIdx) {
    const allKeywords = (typeof MyFamilyProcessor !== 'undefined' && MyFamilyProcessor.getAllCategoryKeywords)
      ? MyFamilyProcessor.getAllCategoryKeywords(profiles)
      : ['HARNESS', 'COLLAR', 'LEASH', '胸背帶', '項圈', '牽繩', '背帶'];
    const maxCol = range ? range.endCell().columnNumber() : 20;
    let bestCol = 0;
    let maxMatch = 0;
    for (let c = 1; c <= maxCol; c++) {
      let matches = 0;
      for (let r = actualStart; r <= Math.min(actualStart + 20, actualEnd); r++) {
        const val = (getCellValue(ws.cell(r, c)) || '').toString().toUpperCase();
        if (allKeywords.some(k => val.includes(k))) matches++;
      }
      if (matches > maxMatch) {
        maxMatch = matches;
        bestCol = c;
      }
    }
    if (bestCol > 0) nameColIdx = bestCol;
  }

  for (let r = actualStart; r <= actualEnd; r++) {
    const firstCell = ws.cell(r, 1);
    const bg = firstCell.style('fill');
    if (bg && (bg.color === 'FAD9D6' || bg === 'FAD9D6' || (typeof bg === 'object' && bg.rgb === 'FAD9D6'))) {
      continue;
    }

    if (filterColIdx) {
      const filterVal = getCellValue(ws.cell(r, filterColIdx));
      if (filterVal === null || filterVal === undefined || String(filterVal).trim() === '') {
        continue;
      }
    }

    let zhName = '';
    if (nameColIdx) {
      zhName = (getCellValue(ws.cell(r, nameColIdx)) || '').toString().trim();
    }
    if (!zhName) continue;

    const rawType = typeColIdx ? (getCellValue(ws.cell(r, typeColIdx)) || '').toString().trim() : '';
    const targetInfo = processor.getTargetTemplateAndCategory(zhName, rawType);
    if (targetInfo && targetInfo.template_id && targetInfo.template_id !== 'UNMATCHED') {
      matchedProfileIds.add(targetInfo.template_id);
    }
  }

  const detected = profiles.filter(p => matchedProfileIds.has(p.id) || matchedProfileIds.has(p.template_type));
  return detected.length > 0 ? detected : profiles;
}

/**
 * Intelligent fuzzy matching to recommend matching source Excel header column
 */
function findBestHeaderSuggestion(targetField, expectedSourceCol, availableHeaders) {
  if (!availableHeaders || availableHeaders.length === 0) return '';
  const cleanStr = (s) => String(s || '').replace(/[\s\r\n\(\)（）\-_*\/\\#\[\]]/g, '').toUpperCase();
  
  const targetClean = cleanStr(targetField);
  const expectedClean = cleanStr(expectedSourceCol);

  // 1. Direct or normalized match
  for (const h of availableHeaders) {
    const hClean = cleanStr(h);
    if (hClean === expectedClean && expectedClean !== '') return h;
  }

  // 2. Synonyms mapping
  const synonyms = [
    { keywords: ['售價', '定價', '訂價', '建議售價', '建議酷澎售價', '台灣訂價', 'PRICE', 'RETAIL'], targets: ['建議酷澎售價', '台灣訂價', '售價'] },
    { keywords: ['進價', '進貨價', '酷澎進價', '成本', 'COST', '採購價', '含稅進價'], targets: ['酷澎進價', '酷澎進價 (含稅)', '進價'] },
    { keywords: ['條碼', 'EAN', 'SKU', 'BARCODE', '商品條碼', '國際條碼', '條碼號', 'GTIN'], targets: ['商品條碼', 'EAN', '條碼'] },
    { keywords: ['品名', '商品名稱', '中文品名', '產品名稱', 'NAME', 'ITEMNAME', '商品名', '中文', '商品'], targets: ['商品名稱', '中文品名', '品名'] },
    { keywords: ['包裝尺寸', '尺寸(MM)', '每單位包裝尺寸', 'DIMENSION', '體積', '規格尺寸'], targets: ['每單位包裝尺寸(mm)', '每單位包裝尺寸\n(mm)', '每單位包裝尺寸'] },
    { keywords: ['包裝重量', '重量(G)', '每單位包裝重量', 'WEIGHT', '重量'], targets: ['每單位包裝重量(g)', '每單位包裝重量\n(g)', '每單位包裝重量'] },
    { keywords: ['背標', '中文背標', '標籤', 'LABEL', 'LABELNAME', '背標檔名'], targets: ['中文背標', '背標'] }
  ];

  for (const group of synonyms) {
    const isTargetMatch = group.targets.some(t => targetClean.includes(cleanStr(t)) || expectedClean.includes(cleanStr(t)));
    if (isTargetMatch) {
      // Pass A: Exact match with keyword
      for (const h of availableHeaders) {
        const hClean = cleanStr(h);
        if (group.keywords.some(k => hClean === cleanStr(k))) {
          return h;
        }
      }
      // Pass B: Substring match (fallback)
      for (const h of availableHeaders) {
        const hClean = cleanStr(h);
        if (group.keywords.some(k => hClean.includes(cleanStr(k)) || cleanStr(k).includes(hClean))) {
          return h;
        }
      }
    }
  }

  // 3. Substring match
  for (const h of availableHeaders) {
    const hClean = cleanStr(h);
    if ((expectedClean && (hClean.includes(expectedClean) || expectedClean.includes(hClean))) ||
        (targetClean && (targetClean.includes(hClean) || hClean.includes(targetClean)))) {
      return h;
    }
  }

  return '';
}

/**
 * Check if the uploaded Excel contains all required columns for the detected template profiles
 */
function checkMissingColumns(detectedProfiles, ws, headerRow, sourceConfig) {
  const headerMap = buildHeaderMap(ws, headerRow);
  const maxCol = ws.usedRange() ? ws.usedRange().endCell().columnNumber() : 50;
  const rawHeaders = [];
  for (let c = 1; c <= maxCol; c++) {
    const val = ws.cell(headerRow, c).value();
    if (val !== null && val !== undefined && String(val).trim() !== '') {
      const raw = String(val).trim();
      if (!rawHeaders.includes(raw)) {
        rawHeaders.push(raw);
      }
    }
  }

  const missingMap = new Map();

  // 1. Check filter column
  const filterCol = sourceConfig?.filter_column || '中文背標';
  const filterIdx = findHeaderColIdx(headerMap, filterCol);
  if (!filterIdx) {
    missingMap.set('__FILTER_COL__', {
      key: '__FILTER_COL__',
      targetField: `篩選欄位（當前設定: ${filterCol}）`,
      expectedSourceCol: filterCol,
      profiles: detectedProfiles,
      isFilter: true
    });
  }

  // 2. Check dynamic fields in detected profiles
  const ignoredSystem = ['細分商品種類', '顏色', '尺寸'];
  for (const p of detectedProfiles) {
    const dyn = p.field_mappings?.dynamic || {};
    for (const [targetKey, expCol] of Object.entries(dyn)) {
      if (ignoredSystem.includes(targetKey)) continue;
      if (!expCol || String(expCol).trim() === '') continue;

      const foundIdx = findHeaderColIdx(headerMap, expCol);
      if (!foundIdx) {
        if (!missingMap.has(targetKey)) {
          missingMap.set(targetKey, {
            key: targetKey,
            targetField: targetKey,
            expectedSourceCol: expCol,
            profiles: [p],
            isFilter: false
          });
        } else {
          const item = missingMap.get(targetKey);
          if (!item.profiles.some(existing => existing.id === p.id)) {
            item.profiles.push(p);
          }
        }
      }
    }
  }

  // 3. Check critical name column
  const nameIdx = findHeaderColIdx(headerMap, '商品名稱') || findHeaderColIdx(headerMap, '中文品名') || findHeaderColIdx(headerMap, '品名') || findHeaderColIdx(headerMap, 'NAME');
  if (!nameIdx && !missingMap.has('商品名稱') && !missingMap.has('中文品名')) {
    missingMap.set('商品名稱', {
      key: '商品名稱',
      targetField: '中文品名 / 商品名稱 (核心必備)',
      expectedSourceCol: '中文品名',
      profiles: detectedProfiles,
      isFilter: false
    });
  }

  const missingItems = Array.from(missingMap.values()).map(item => {
    const suggestion = findBestHeaderSuggestion(item.targetField, item.expectedSourceCol, rawHeaders);
    return {
      ...item,
      suggestion: suggestion
    };
  });

  return {
    detectedProfiles,
    missingItems,
    excelHeaderNames: rawHeaders
  };
}

document.addEventListener('DOMContentLoaded', async () => {
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
  
  const profileListContainer = document.getElementById('profileListContainer');
  const profileEditorContent = document.getElementById('profileEditorContent');
  const btnOpenUploadWizard = document.getElementById('btnOpenUploadWizard');
  const uploadTemplateFileInput = document.getElementById('uploadTemplateFileInput');

  const guiSourceConfigForm = document.getElementById('guiSourceConfigForm');
  const guiCollectionForm = document.getElementById('guiCollectionForm');
  const btnAddCollection = document.getElementById('btnAddCollection');
  const guiColorForm = document.getElementById('guiColorForm');
  const btnAddColor = document.getElementById('btnAddColor');
  const btnRescanCollection = document.getElementById('btnRescanCollection');
  const btnRescanColor = document.getElementById('btnRescanColor');

  const btnExportConfigPackage = document.getElementById('btnExportConfigPackage');
  const btnImportConfigPackage = document.getElementById('btnImportConfigPackage');
  const importConfigFileInput = document.getElementById('importConfigFileInput');

  // Wizard elements
  const templateWizardModal = document.getElementById('templateWizardModal');
  const wizardModalClose = document.getElementById('wizardModalClose');
  const btnCancelWizard = document.getElementById('btnCancelWizard');
  const btnConfirmWizard = document.getElementById('btnConfirmWizard');
  const wizardProfileName = document.getElementById('wizardProfileName');
  const wizardCategoryName = document.getElementById('wizardCategoryName');
  const wizardKeywords = document.getElementById('wizardKeywords');
  const wizardSubfolder = document.getElementById('wizardSubfolder');
  const wizardScannedCount = document.getElementById('wizardScannedCount');
  const wizardTableBody = document.getElementById('wizardTableBody');

  // Column Mapping Modal elements
  const columnMappingModal = document.getElementById('columnMappingModal');
  const mappingModalClose = document.getElementById('mappingModalClose');
  const detectedTemplatesContainer = document.getElementById('detectedTemplatesContainer');
  const missingColumnCount = document.getElementById('missingColumnCount');
  const mappingTableBody = document.getElementById('mappingTableBody');
  const chkRememberMappings = document.getElementById('chkRememberMappings');
  const btnCancelMapping = document.getElementById('btnCancelMapping');
  const btnConfirmMapping = document.getElementById('btnConfirmMapping');

  // State Variables
  let templateProfiles = [];
  let activeProfileId = 'HARNESS';
  let pendingWizard = null;

  let currentSourceConfig = window.AppConfig.get().source;
  let currentCollectionAliases = window.AppConfig.getCollectionAliases();
  let currentColorAliases = window.AppConfig.getColorAliases();

  let sourceExcelFile = null;
  let loadedWorkbook = null;
  let photoFilesArray = [];
  let processedResults = [];
  let generatedExcelFiles = new Map(); // Map<subfolder, Map<filename, buffer>>
  let filesToExport = {};

  // Initialize Template Profiles from Storage / IndexedDB
  async function initTemplateProfiles() {
    try {
      const stored = await window.StorageUtils.getAllProfiles();
      if (Array.isArray(stored) && stored.length > 0) {
        // Auto-migrate if stored contains old merged '項圈牽繩' profile
        const hasMerged = stored.some(p => p.id === 'LEASH' && (p.name === '項圈牽繩' || (p.keywords && p.keywords.includes('項圈') && p.keywords.includes('牽繩'))));
        if (hasMerged && !stored.some(p => p.id === 'COLLAR')) {
          const defaults = window.AppConfig.getDefaultProfiles();
          const customProfiles = stored.filter(p => !p.is_builtin);
          templateProfiles = [...defaults, ...customProfiles];
          for (const p of templateProfiles) {
            await window.StorageUtils.saveProfile(p);
          }
        } else {
          templateProfiles = stored;
        }
      } else {
        templateProfiles = window.AppConfig.getDefaultProfiles();
        for (const p of templateProfiles) {
          await window.StorageUtils.saveProfile(p);
        }
      }
    } catch (e) {
      console.warn('Failed to load profiles from IndexedDB, using defaults:', e);
      templateProfiles = window.AppConfig.getDefaultProfiles();
    }
  }

  await initTemplateProfiles();

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

  // Render Template Profiles Sidebar and Editor
  function renderTemplateProfilesUI() {
    if (!profileListContainer) return;
    profileListContainer.innerHTML = '';

    templateProfiles.forEach(p => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `profile-item ${p.id === activeProfileId ? 'active' : ''}`;
      btn.innerHTML = `
        <span>${p.name || '未命名'}</span>
        <span class="profile-item-badge">${p.is_builtin ? '預設' : '自訂'}</span>
      `;
      btn.addEventListener('click', () => {
        saveActiveProfileFromUI();
        activeProfileId = p.id;
        renderTemplateProfilesUI();
      });
      profileListContainer.appendChild(btn);
    });

    renderActiveProfileEditor();
  }

  function renderActiveProfileEditor() {
    if (!profileEditorContent) return;
    const profile = templateProfiles.find(p => p.id === activeProfileId) || templateProfiles[0];
    if (!profile) {
      profileEditorContent.innerHTML = '<p style="color:#64748b; padding:20px;">請先由左側選擇或新增設定檔。</p>';
      return;
    }
    activeProfileId = profile.id;

    let fixedHtml = '';
    for (const [k, v] of Object.entries(profile.field_mappings?.fixed || {})) {
      fixedHtml += `
        <div class="dynamic-row fixed-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${k}">
          <input type="text" class="val-input" placeholder="固定填寫內容" value="${v}">
          <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-icons">delete</span></button>
        </div>`;
    }

    let dynamicHtml = '';
    for (const [k, v] of Object.entries(profile.field_mappings?.dynamic || {})) {
      dynamicHtml += `
        <div class="dynamic-row dynamic-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${k}">
          <input type="text" class="val-input" placeholder="來源表對應欄位" value="${v}">
          <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-icons">delete</span></button>
        </div>`;
    }

    const keywordsStr = Array.isArray(profile.keywords) ? profile.keywords.join(', ') : (profile.keywords || profile.name || '');

    profileEditorContent.innerHTML = `
      <div class="profile-meta-card">
        <div class="profile-meta-grid">
          <div class="input-row" style="margin:0;">
            <label style="flex:0 0 110px; font-weight:600;">設定檔名稱:</label>
            <input type="text" id="prof_name" value="${profile.name || ''}" ${profile.is_builtin ? 'readonly style="background:#f1f5f9;"' : ''}>
          </div>
          <div class="input-row" style="margin:0;">
            <label style="flex:0 0 110px; font-weight:600;">輸出子資料夾:</label>
            <input type="text" id="prof_subfolder" value="${profile.subfolder || profile.name || ''}">
          </div>
          <div class="input-row" style="margin:0;">
            <label style="flex:0 0 110px; font-weight:600;">匹配關鍵字:</label>
            <input type="text" id="prof_keywords" placeholder="例: HARNESS, 胸背帶 (逗號分隔)" value="${keywordsStr}">
          </div>
          <div class="input-row" style="margin:0;">
            <label style="flex:0 0 110px; font-weight:600;">預設分類代碼:</label>
            <input type="text" id="prof_category_name" placeholder="例如: 寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)" value="${profile.category_name || ''}">
          </div>
        </div>
        <div class="template-file-bar" style="margin-top:12px;">
          <div class="template-file-info">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
            <span>Excel 範本: <strong>${profile.template_file_name || (profile.id + '.xlsx')}</strong></span>
            <span class="badge-${profile.is_builtin ? 'system' : 'inherited'}">${profile.is_builtin ? '內建範本' : '自訂上傳'}</span>
          </div>
          <div class="template-file-actions">
            <button type="button" id="btnDownloadProfileTemplate" class="btn btn-outline btn-sm btn-icon" title="下載 Excel 範本檔"><span class="material-icons">download</span></button>
            <button type="button" id="btnReplaceProfileTemplate" class="btn btn-secondary btn-sm btn-icon" title="替換 Excel 範本檔"><span class="material-icons">file_upload</span></button>
            <input type="file" id="replaceTemplateFileInput" accept=".xlsx" hidden>
          </div>
        </div>
      </div>

      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 1px solid #e2e8f0; margin-bottom:10px; padding-bottom:6px;">
          <h4 style="border:none; margin:0; padding:0;">固定欄位對應 (Fixed Mappings)</h4>
          <button type="button" class="btn btn-outline btn-sm btn-icon" title="新增固定欄位" onclick="addProfileFixedRow()"><span class="material-icons">add</span></button>
        </div>
        <p style="font-size:0.82rem; color:#64748b; margin-bottom:8px;">不論來源資料為何，強制填入目標 Excel 模板的固定值。</p>
        <div id="profileFixedContainer">${fixedHtml}</div>
      </div>

      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 1px solid #e2e8f0; margin-bottom:10px; padding-bottom:6px;">
          <h4 style="border:none; margin:0; padding:0;">動態欄位對應 (Dynamic Mappings)</h4>
          <button type="button" class="btn btn-outline btn-sm btn-icon" title="新增動態欄位" onclick="addProfileDynamicRow()"><span class="material-icons">add</span></button>
        </div>
        <p style="font-size:0.82rem; color:#64748b; margin-bottom:8px;">將來源商品表的欄位資料，動態填入目標 Excel 模板的對應欄位中。</p>
        <div id="profileDynamicContainer">${dynamicHtml}</div>
      </div>

      ${!profile.is_builtin ? `
        <div style="display:flex; justify-content:flex-end; margin-top:20px; padding-top:10px; border-top:1px dashed #e2e8f0;">
          <button type="button" id="btnDeleteCurrentProfile" class="btn btn-danger btn-sm btn-icon" title="刪除此模板設定檔"><span class="material-icons">delete_forever</span></button>
        </div>
      ` : ''}
    `;

    // Hook template actions
    const btnDownloadProfileTemplate = document.getElementById('btnDownloadProfileTemplate');
    if (btnDownloadProfileTemplate) {
      btnDownloadProfileTemplate.addEventListener('click', async () => {
        let b64 = null;
        if (profile.is_builtin) {
          b64 = (profile.id === 'HARNESS') ? window.CoupangTemplates?.HARNESS : window.CoupangTemplates?.LEASH;
        }
        if (!b64) {
          b64 = await window.StorageUtils.getTemplateData(profile.id);
        }
        if (!b64) {
          alert('找不到該模板檔案！');
          return;
        }
        const ab = b64ToArrayBuffer(b64);
        const blob = new Blob([ab], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        saveAs(blob, profile.template_file_name || `${profile.name}.xlsx`);
      });
    }

    const btnReplaceProfileTemplate = document.getElementById('btnReplaceProfileTemplate');
    const replaceTemplateFileInput = document.getElementById('replaceTemplateFileInput');
    if (btnReplaceProfileTemplate && replaceTemplateFileInput) {
      btnReplaceProfileTemplate.addEventListener('click', () => replaceTemplateFileInput.click());
      replaceTemplateFileInput.addEventListener('change', async (e) => {
        if (e.target.files.length > 0) {
          const file = e.target.files[0];
          const b64 = await readAsBase64(file);
          profile.template_file_name = file.name;
          await window.StorageUtils.saveTemplateData(profile.id, b64);
          await window.StorageUtils.saveProfile(profile);
          alert(`已成功替換【${profile.name}】的 Excel 範本！`);
          renderActiveProfileEditor();
        }
      });
    }

    const btnDeleteCurrentProfile = document.getElementById('btnDeleteCurrentProfile');
    if (btnDeleteCurrentProfile) {
      btnDeleteCurrentProfile.addEventListener('click', async () => {
        if (confirm(`確定要刪除設定檔【${profile.name}】嗎？`)) {
          await window.StorageUtils.deleteProfile(profile.id);
          templateProfiles = templateProfiles.filter(p => p.id !== profile.id);
          activeProfileId = templateProfiles[0]?.id || 'HARNESS';
          renderTemplateProfilesUI();
        }
      });
    }
  }

  window.addProfileFixedRow = function(k='', v='') {
    const div = document.createElement('div');
    div.className = 'dynamic-row fixed-field-row';
    div.innerHTML = `
      <input type="text" class="key-input" placeholder="目標模板欄位" value="${k}">
      <input type="text" class="val-input" placeholder="固定填寫內容" value="${v}">
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-icons">delete</span></button>
    `;
    document.getElementById('profileFixedContainer')?.appendChild(div);
  };

  window.addProfileDynamicRow = function(k='', v='') {
    const div = document.createElement('div');
    div.className = 'dynamic-row dynamic-field-row';
    div.innerHTML = `
      <input type="text" class="key-input" placeholder="目標模板欄位" value="${k}">
      <input type="text" class="val-input" placeholder="來源表對應欄位" value="${v}">
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-icons">delete</span></button>
    `;
    document.getElementById('profileDynamicContainer')?.appendChild(div);
  };

  function saveActiveProfileFromUI() {
    const profile = templateProfiles.find(p => p.id === activeProfileId);
    if (!profile) return;

    const nameInput = document.getElementById('prof_name');
    const subfolderInput = document.getElementById('prof_subfolder');
    const categoryInput = document.getElementById('prof_category_name');
    const keywordsInput = document.getElementById('prof_keywords');

    if (nameInput && !profile.is_builtin) profile.name = nameInput.value.trim();
    if (subfolderInput) profile.subfolder = subfolderInput.value.trim() || profile.name;
    if (categoryInput) profile.category_name = categoryInput.value.trim();
    if (keywordsInput) {
      profile.keywords = keywordsInput.value.split(/[,，]/).map(k => k.trim()).filter(Boolean);
    }

    const fixed = {};
    document.querySelectorAll('#profileFixedContainer .fixed-field-row').forEach(row => {
      const k = row.querySelector('.key-input').value.trim();
      const v = row.querySelector('.val-input').value.trim();
      if (k) fixed[k] = v;
    });

    const dynamic = {};
    document.querySelectorAll('#profileDynamicContainer .dynamic-field-row').forEach(row => {
      const k = row.querySelector('.key-input').value.trim();
      const v = row.querySelector('.val-input').value.trim();
      if (k) dynamic[k] = v;
    });

    profile.field_mappings = { fixed, dynamic };
  }

  // Render Source Config
  function renderSourceConfig(source) {
    if (!guiSourceConfigForm) return;
    guiSourceConfigForm.innerHTML = `
      <div class="form-group">
        <h4>來源表解析設定 (Source Settings)</h4>
        <div class="input-row"><label>來源表名稱:</label><input type="text" id="cfg_sheet_name" value="${source.sheet_name || ''}"></div>
        <div class="input-row"><label>標題列位於第幾列:</label><input type="number" id="cfg_header_row" value="${source.header_row || 3}"></div>
        <div class="input-row"><label>資料起始列:</label><input type="number" id="cfg_row_start" value="${source.row_start || 4}"></div>
        <div class="input-row"><label>篩選欄位(有填寫才處理):</label><input type="text" id="cfg_filter_column" value="${source.filter_column || ''}"></div>
      </div>
    `;
  }

  function parseSourceConfig() {
    return {
      file_path: "商品資料.xlsx",
      sheet_name: document.getElementById('cfg_sheet_name')?.value.trim() || 'Sheet1',
      header_row: parseInt(document.getElementById('cfg_header_row')?.value) || 3,
      row_start: parseInt(document.getElementById('cfg_row_start')?.value) || 4,
      filter_column: document.getElementById('cfg_filter_column')?.value.trim() || '中文背標'
    };
  }

  function renderGuiCollection(collectionAliases) {
    if (!guiCollectionForm) return;
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
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此系列對照" onclick="this.parentElement.remove()"><span class="material-icons">delete</span></button>
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
    if (!guiColorForm) return;
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
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此顏色群組" onclick="this.parentElement.remove()"><span class="material-icons">delete</span></button>
    `;
    guiColorForm.appendChild(row);
  }

  if (btnAddColor) {
    btnAddColor.addEventListener('click', () => addColorRow());
  }

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

  // Wizard Launch & Confirmation
  if (btnOpenUploadWizard && uploadTemplateFileInput) {
    btnOpenUploadWizard.addEventListener('click', () => {
      uploadTemplateFileInput.click();
    });

    uploadTemplateFileInput.addEventListener('change', async (e) => {
      if (e.target.files.length > 0) {
        const file = e.target.files[0];
        try {
          const ab = await file.arrayBuffer();
          const scanResult = await scanTemplateHeaders(ab);
          const base64 = await readAsBase64(file);
          pendingWizard = { file, ab, base64, scanResult };

          const defaultName = file.name.replace(/\.xlsx$/i, '').replace(/^商品報價單_/, '');
          wizardProfileName.value = defaultName;
          wizardCategoryName.value = '';
          wizardSubfolder.value = defaultName;
          wizardScannedCount.textContent = scanResult.requiredColumns.length;

          wizardTableBody.innerHTML = '';
          scanResult.requiredColumns.forEach((col, idx) => {
            const tr = document.createElement('tr');
            let statusBadge = '';
            if (col.status === 'system') {
              statusBadge = '<span class="badge-system">系統自動處理</span>';
            } else if (col.status === 'inherited') {
              statusBadge = '<span class="badge-inherited">繼承自預設</span>';
            } else {
              statusBadge = '<span class="badge-new">新必填欄位</span>';
            }

            const isSys = col.status === 'system';
            tr.innerHTML = `
              <td><strong>${col.colName}</strong></td>
              <td>${statusBadge}</td>
              <td>
                <select class="wizard-type-select" data-idx="${idx}" ${isSys ? 'disabled' : ''}>
                  <option value="dynamic" ${col.mappingType === 'dynamic' ? 'selected' : ''}>動態對應 (來源表)</option>
                  <option value="fixed" ${col.mappingType === 'fixed' ? 'selected' : ''}>固定值 (固定內容)</option>
                  <option value="system" ${col.mappingType === 'system' ? 'selected' : ''}>系統自動處理</option>
                </select>
              </td>
              <td>
                <input type="text" class="wizard-val-input" data-idx="${idx}" value="${col.defaultValue}" placeholder="${col.mappingType === 'dynamic' ? '來源表欄位名 (例: 中文品名)' : '固定填寫內容'}" ${isSys ? 'disabled' : ''}>
              </td>
            `;

            const select = tr.querySelector('.wizard-type-select');
            const input = tr.querySelector('.wizard-val-input');
            select.addEventListener('change', () => {
              if (select.value === 'dynamic') {
                input.disabled = false;
                input.placeholder = '來源表欄位名 (例: 中文品名)';
              } else if (select.value === 'fixed') {
                input.disabled = false;
                input.placeholder = '固定填寫內容';
              } else {
                input.disabled = true;
                input.value = '(系統自動處理)';
              }
            });

            wizardTableBody.appendChild(tr);
          });

          templateWizardModal.classList.remove('hidden');
        } catch (err) {
          console.error('掃描 Excel 模板失敗:', err);
          alert('掃描 Excel 模板失敗: ' + err.message);
        } finally {
          uploadTemplateFileInput.value = '';
        }
      }
    });
  }

  if (wizardModalClose) wizardModalClose.addEventListener('click', () => templateWizardModal.classList.add('hidden'));
  if (btnCancelWizard) btnCancelWizard.addEventListener('click', () => templateWizardModal.classList.add('hidden'));

  if (btnConfirmWizard) {
    btnConfirmWizard.addEventListener('click', async () => {
      if (!pendingWizard) return;
      const profName = wizardProfileName.value.trim();
      if (!profName) {
        alert('請輸入設定檔名稱！');
        return;
      }

      const categoryName = wizardCategoryName.value.trim();
      const subfolder = wizardSubfolder.value.trim() || profName;

      const dynamic = {};
      const fixed = {};

      const rows = wizardTableBody.querySelectorAll('tr');
      rows.forEach((tr, idx) => {
        const col = pendingWizard.scanResult.requiredColumns[idx];
        const typeSelect = tr.querySelector('.wizard-type-select');
        const valInput = tr.querySelector('.wizard-val-input');
        const mType = typeSelect.value;
        const val = valInput.value.trim();

        if (mType === 'dynamic' && val) {
          dynamic[col.colName] = val;
        } else if (mType === 'fixed' && val) {
          fixed[col.colName] = val;
        }
      });

      const profileId = 'profile_' + Date.now();
      const kwStr = (wizardKeywords ? wizardKeywords.value : '').trim();
      const keywords = kwStr ? kwStr.split(/[,，]/).map(k => k.trim()).filter(Boolean) : [profName];

      const newProfile = {
        id: profileId,
        name: profName,
        keywords: keywords,
        template_type: profileId,
        template_id: profileId,
        template_file_name: pendingWizard.file.name,
        category_name: categoryName,
        subfolder: subfolder,
        is_builtin: false,
        field_mappings: { dynamic, fixed }
      };

      try {
        await window.StorageUtils.saveTemplateData(profileId, pendingWizard.base64);
        await window.StorageUtils.saveProfile(newProfile);
        templateProfiles.push(newProfile);

        activeProfileId = profileId;
        renderTemplateProfilesUI();

        templateWizardModal.classList.add('hidden');
        alert(`已成功建立【${profName}】模板設定檔！`);
      } catch (err) {
        console.error('儲存新設定檔失敗:', err);
        alert('儲存新設定檔失敗: ' + err.message);
      }
    });
  }

  // Column Mapping Modal for Missing Fields
  function showColumnMappingModal(detectedProfiles, missingItems, excelHeaderNames) {
    return new Promise((resolve) => {
      if (!columnMappingModal) {
        resolve({ action: 'cancel' });
        return;
      }

      // Render detected templates badges
      if (detectedTemplatesContainer) {
        detectedTemplatesContainer.innerHTML = '';
        detectedProfiles.forEach(p => {
          const badge = document.createElement('span');
          badge.className = 'badge-template';
          badge.textContent = p.name;
          detectedTemplatesContainer.appendChild(badge);
        });
      }

      // Render count
      if (missingColumnCount) {
        missingColumnCount.textContent = missingItems.length;
      }

      // Render table rows
      if (mappingTableBody) {
        mappingTableBody.innerHTML = '';
        missingItems.forEach(item => {
          const tr = document.createElement('tr');

          const tdTarget = document.createElement('td');
          tdTarget.innerHTML = `<span class="mapping-target-name">${item.targetField}</span>`;

          const tdProfiles = document.createElement('td');
          tdProfiles.textContent = item.profiles.map(p => p.name).join('、');

          const tdExpected = document.createElement('td');
          tdExpected.innerHTML = `<span class="mapping-expected-name">${item.expectedSourceCol}</span>`;

          const tdSelect = document.createElement('td');
          const box = document.createElement('div');
          box.className = 'mapping-select-box';

          const select = document.createElement('select');
          select.className = 'mapping-col-select';
          select.dataset.key = item.key;

          const defaultOpt = document.createElement('option');
          defaultOpt.value = '';
          defaultOpt.textContent = '-- 請選擇對應的來源 Excel 欄位 --';
          select.appendChild(defaultOpt);

          excelHeaderNames.forEach(h => {
            const opt = document.createElement('option');
            opt.value = h;
            opt.textContent = (h === item.suggestion) ? `${h} (推薦比對)` : h;
            if (h === item.suggestion) {
              opt.selected = true;
            }
            select.appendChild(opt);
          });

          if (!item.isFilter) {
            const skipOpt = document.createElement('option');
            skipOpt.value = '__SKIP__';
            skipOpt.textContent = '【留空 / 略過此欄位】';
            select.appendChild(skipOpt);

            const fixedOpt = document.createElement('option');
            fixedOpt.value = '__FIXED__';
            fixedOpt.textContent = '【手動輸入固定值...】';
            select.appendChild(fixedOpt);
          }

          box.appendChild(select);

          const fixedInput = document.createElement('input');
          fixedInput.type = 'text';
          fixedInput.className = 'mapping-fixed-input hidden';
          fixedInput.placeholder = `請輸入「${item.targetField}」的固定值內容...`;
          box.appendChild(fixedInput);

          select.addEventListener('change', () => {
            if (select.value === '__FIXED__') {
              fixedInput.classList.remove('hidden');
              fixedInput.focus();
            } else {
              fixedInput.classList.add('hidden');
            }
          });

          tdSelect.appendChild(box);

          tr.appendChild(tdTarget);
          tr.appendChild(tdProfiles);
          tr.appendChild(tdExpected);
          tr.appendChild(tdSelect);
          mappingTableBody.appendChild(tr);
        });
      }

      columnMappingModal.classList.remove('hidden');

      function cleanup() {
        columnMappingModal.classList.add('hidden');
        btnConfirmMapping.removeEventListener('click', onConfirm);
        btnCancelMapping.removeEventListener('click', onCancel);
        mappingModalClose.removeEventListener('click', onCancel);
      }

      function onConfirm() {
        const mappings = {};
        const rows = mappingTableBody.querySelectorAll('tr');
        let hasUnselected = false;

        rows.forEach((r, idx) => {
          const item = missingItems[idx];
          const select = r.querySelector('.mapping-col-select');
          const fixedInput = r.querySelector('.mapping-fixed-input');
          const val = select.value;

          if (!val) {
            hasUnselected = true;
            select.style.borderColor = '#ef4444';
          } else {
            select.style.borderColor = '';
          }

          if (val === '__SKIP__') {
            mappings[item.key] = { type: 'skip', value: '', isFilter: item.isFilter, profiles: item.profiles };
          } else if (val === '__FIXED__') {
            mappings[item.key] = { type: 'fixed', value: (fixedInput.value || '').trim(), isFilter: item.isFilter, profiles: item.profiles };
          } else {
            mappings[item.key] = { type: 'column', value: val, isFilter: item.isFilter, profiles: item.profiles };
          }
        });

        if (hasUnselected) {
          alert('請為所有缺漏欄位選取對應來源、或選擇「留空」/「手動輸入固定值」！');
          return;
        }

        const remember = chkRememberMappings ? chkRememberMappings.checked : true;
        cleanup();
        resolve({ action: 'confirm', mappings, remember });
      }

      function onCancel() {
        cleanup();
        resolve({ action: 'cancel' });
      }

      btnConfirmMapping.addEventListener('click', onConfirm);
      btnCancelMapping.addEventListener('click', onCancel);
      mappingModalClose.addEventListener('click', onCancel);
    });
  }

  // Backup Export and Import
  if (btnExportConfigPackage) {
    btnExportConfigPackage.addEventListener('click', async () => {
      try {
        saveActiveProfileFromUI();
        if (Array.isArray(templateProfiles)) {
          for (const p of templateProfiles) {
            await window.StorageUtils.saveProfile(p);
          }
        }
        if (typeof parseSourceConfig === 'function') {
          const sc = parseSourceConfig();
          const globalConfig = window.AppConfig.get();
          globalConfig.source = sc;
          localStorage.setItem('coupang_config', JSON.stringify(globalConfig));
          localStorage.setItem('my_family_config', JSON.stringify(globalConfig));
        }
        if (typeof parseGuiCollection === 'function') {
          const ca = parseGuiCollection();
          localStorage.setItem('coupang_collection_aliases', JSON.stringify(ca));
          localStorage.setItem('my_family_collection_aliases', JSON.stringify(ca));
        }
        if (typeof parseGuiColor === 'function') {
          const cla = parseGuiColor();
          localStorage.setItem('coupang_color_aliases', JSON.stringify(cla));
          localStorage.setItem('my_family_color_aliases', JSON.stringify(cla));
        }
        const pkg = await window.StorageUtils.exportConfigPackage();
        const jsonStr = JSON.stringify(pkg, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        saveAs(blob, `coupang_config_backup_${new Date().toISOString().slice(0, 10)}.json`);
        logMessage('設定檔備份 JSON 已成功匯出！', 'success');
      } catch (err) {
        alert('匯出設定檔備份失敗: ' + err.message);
      }
    });
  }

  if (btnImportConfigPackage && importConfigFileInput) {
    btnImportConfigPackage.addEventListener('click', () => importConfigFileInput.click());
    importConfigFileInput.addEventListener('change', async (e) => {
      if (e.target.files.length > 0) {
        try {
          const file = e.target.files[0];
          const text = await file.text();
          const pkg = JSON.parse(text);
          await window.StorageUtils.importConfigPackage(pkg);
          await initTemplateProfiles();
          currentSourceConfig = window.AppConfig.get().source;
          currentCollectionAliases = window.AppConfig.getCollectionAliases();
          currentColorAliases = window.AppConfig.getColorAliases();
          
          renderTemplateProfilesUI();
          renderSourceConfig(currentSourceConfig);
          renderGuiCollection(currentCollectionAliases);
          renderGuiColor(currentColorAliases);

          alert('設定檔與模板已成功匯入還原！');
          logMessage('已從備份檔完整還原所有設定檔與模板！', 'success');
        } catch (err) {
          console.error('匯入設定失敗:', err);
          alert('匯入設定失敗: ' + err.message);
        } finally {
          importConfigFileInput.value = '';
        }
      }
    });
  }

  // Modal Open & Tab Switching
  btnOpenConfig.addEventListener('click', () => {
    renderTemplateProfilesUI();
    renderSourceConfig(currentSourceConfig);
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
      document.getElementById(btn.dataset.target)?.classList.add('active');
    });
  });

  // Save and Reset in Config Modal
  btnSaveConfig.addEventListener('click', async () => {
    try {
      saveActiveProfileFromUI();
      for (const p of templateProfiles) {
        await window.StorageUtils.saveProfile(p);
      }

      currentSourceConfig = parseSourceConfig();
      currentCollectionAliases = parseGuiCollection();
      currentColorAliases = parseGuiColor();

      const globalConfig = window.AppConfig.get();
      globalConfig.source = currentSourceConfig;
      localStorage.setItem('coupang_config', JSON.stringify(globalConfig));
      localStorage.setItem('my_family_config', JSON.stringify(globalConfig));

      localStorage.setItem('coupang_collection_aliases', JSON.stringify(currentCollectionAliases));
      localStorage.setItem('my_family_collection_aliases', JSON.stringify(currentCollectionAliases));

      localStorage.setItem('coupang_color_aliases', JSON.stringify(currentColorAliases));
      localStorage.setItem('my_family_color_aliases', JSON.stringify(currentColorAliases));

      updateRangeHintUI();
      alert('所有設定檔與對照表已儲存！');
      configModal.classList.add('hidden');
    } catch (e) {
      alert('儲存設定時發生錯誤: ' + e.message);
    }
  });

  btnResetConfig.addEventListener('click', async () => {
    if (confirm('確定要還原為系統預設值嗎？自訂模板設定檔將被清除。')) {
      localStorage.removeItem('coupang_config');
      localStorage.removeItem('coupang_category_rules');
      localStorage.removeItem('coupang_collection_aliases');
      localStorage.removeItem('coupang_color_aliases');
      localStorage.removeItem('coupang_templates');
      localStorage.removeItem('coupang_template_profiles');

      templateProfiles = window.AppConfig.getDefaultProfiles();
      for (const p of templateProfiles) {
        await window.StorageUtils.saveProfile(p);
      }

      currentSourceConfig = window.AppConfig.getDefaultSourceConfig();
      currentCollectionAliases = window.AppConfig.getDefaultCollectionAliases();
      currentColorAliases = window.AppConfig.getDefaultColorAliases();

      renderTemplateProfilesUI();
      renderSourceConfig(currentSourceConfig);
      renderGuiCollection(currentCollectionAliases);
      renderGuiColor(currentColorAliases);

      if (inputRowStart) inputRowStart.value = currentSourceConfig.row_start || 4;
      if (inputRowEnd) inputRowEnd.value = "";
      updateRangeHintUI();
      alert('已還原為系統預設值。');
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

  async function autoScanAndSyncMappings(showLog = true) {
    let scannedFolder = { collections: [], colors: [] };
    if (photoFilesArray && photoFilesArray.length > 0) {
      scannedFolder = window.MyFamilyProcessor.scanFolderStructure(photoFilesArray, templateProfiles);
    }

    let excelData = { collections: {}, colors: {} };
    if (loadedWorkbook) {
      const headerRow = currentSourceConfig.header_row || 3;
      const rowStart = currentSourceConfig.row_start || 4;
      const filterColName = currentSourceConfig.filter_column || '中文背標';
      excelData = window.MyFamilyProcessor.extractMappingsFromExcel(loadedWorkbook, headerRow, rowStart, filterColName, currentSourceConfig?.sheet_name, templateProfiles);
    }

    if (scannedFolder.collections.length > 0 || Object.keys(excelData.collections).length > 0 ||
        scannedFolder.colors.length > 0 || Object.keys(excelData.colors).length > 0) {
      
      const merged = window.MyFamilyProcessor.mergeScannedAliases(
        currentCollectionAliases,
        currentColorAliases,
        scannedFolder,
        excelData,
        templateProfiles
      );

      currentCollectionAliases = merged.collectionAliases;
      currentColorAliases = merged.colorAliases;

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

  // Drag & drop handlers
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

  function getSourceSheet(wb) {
    if (!wb) return null;
    const headerRow = currentSourceConfig?.header_row || 3;
    const targetSheetName = currentSourceConfig?.sheet_name;

    if (targetSheetName) {
      const s = wb.sheet(targetSheetName);
      if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
        return s;
      }
    }

    const candidateNames = ['商品資料', 'MYFAMILY', 'MY FAMILY', 'My Family', '工作表1', 'Sheet1'];
    for (const name of candidateNames) {
      const s = wb.sheet(name);
      if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
        return s;
      }
    }

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
    if (!ws) return { validRows: [], totalDataRows: 0, startRow: 0, endRow: 0 };
    const headerRow = currentSourceConfig?.header_row || 3;
    const defaultStartRow = currentSourceConfig?.row_start || 4;
    const filterColName = currentSourceConfig?.filter_column || '中文背標';
    
    const range = ws.usedRange();
    if (!range) return { validRows: [], totalDataRows: 0, startRow: 0, endRow: 0 };
    const maxRow = range.endCell().rowNumber();

    const headers = buildHeaderMap(ws, headerRow);
    const filterColIdx = findHeaderColIdx(headers, filterColName);

    const actualStart = Math.max(headerRow + 1, parseInt(customStart) || defaultStartRow);
    const actualEnd = customEnd ? Math.min(maxRow, parseInt(customEnd)) : maxRow;

    const validRows = [];
    let totalDataRows = 0;

    for (let r = actualStart; r <= actualEnd; r++) {
      totalDataRows++;
      
      const firstCell = ws.cell(r, 1);
      const bg = firstCell.style('fill');
      if (bg && (bg.color === 'FAD9D6' || bg === 'FAD9D6' || (typeof bg === 'object' && bg.rgb === 'FAD9D6'))) {
        continue;
      }

      if (filterColIdx) {
        const filterVal = getCellValue(ws.cell(r, filterColIdx));
        if (filterVal === null || filterVal === undefined || String(filterVal).trim() === '') {
          continue;
        }
      }

      validRows.push(r);
    }

    return { validRows, totalDataRows, startRow: actualStart, endRow: actualEnd, maxRow };
  }

  function calculateValidProducts(ws) {
    const sStart = inputRowStart ? inputRowStart.value : null;
    const sEnd = inputRowEnd ? inputRowEnd.value : null;
    const info = getValidRowsInfo(ws, sStart, sEnd);
    return { count: info.validRows.length, rows: info.validRows, totalDataRows: info.totalDataRows, maxRow: info.maxRow };
  }

  function updateRangeHintUI() {
    if (!rangeInfoHint) return;
    if (!loadedWorkbook) {
      rangeInfoHint.textContent = '尚未載入來源 Excel';
      rangeInfoHint.className = 'range-hint';
      return;
    }
    const ws = getSourceSheet(loadedWorkbook);
    if (!ws) {
      rangeInfoHint.textContent = '來源 Excel 無有效工作表';
      rangeInfoHint.className = 'range-hint warning';
      return;
    }
    const sStart = inputRowStart ? inputRowStart.value : null;
    const sEnd = inputRowEnd ? inputRowEnd.value : null;
    const info = getValidRowsInfo(ws, sStart, sEnd);
    rangeInfoHint.textContent = `掃描範圍: 第 ${info.startRow} ~ ${info.endRow} 列（總共 ${info.totalDataRows} 列，符合背標條件之有效商品: ${info.validRows.length} 筆）`;
    rangeInfoHint.className = 'range-hint success';
  }

  if (inputRowStart) inputRowStart.addEventListener('input', () => updateRangeHintUI());
  if (inputRowEnd) inputRowEnd.addEventListener('input', () => updateRangeHintUI());
  if (btnResetRange) {
    btnResetRange.addEventListener('click', () => {
      if (inputRowStart) inputRowStart.value = currentSourceConfig.row_start || 4;
      if (inputRowEnd) inputRowEnd.value = '';
      updateRangeHintUI();
    });
  }

  async function handleExcelFile(file) {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const wb = await XlsxPopulate.fromDataAsync(arrayBuffer);
      const ws = getSourceSheet(wb);
      if (!ws || !ws.usedRange()) {
        throw new Error('來源 Excel 中找不到任何可用的工作表！');
      }

      const headerRow = currentSourceConfig?.header_row || 3;
      const totalR = ws.usedRange().endCell().rowNumber();

      // Step 1: Automatically detect required templates
      const detectedProfiles = detectRequiredTemplates(ws, headerRow, currentSourceConfig?.row_start || 4, totalR, templateProfiles, currentSourceConfig);

      // Step 2: Check for missing columns
      const checkResult = checkMissingColumns(detectedProfiles, ws, headerRow, currentSourceConfig);

      // Step 3: Prompt user if missing columns are found
      if (checkResult.missingItems && checkResult.missingItems.length > 0) {
        const modalRes = await showColumnMappingModal(detectedProfiles, checkResult.missingItems, checkResult.excelHeaderNames);
        if (modalRes.action === 'cancel') {
          sourceExcelFile = null;
          loadedWorkbook = null;
          excelDropZone.classList.remove('has-file');
          excelFileInfo.classList.add('hidden');
          excelFileInfo.textContent = '';
          checkReady();
          updateRangeHintUI();
          logMessage('已取消載入來源 Excel 檔案。', 'warning');
          return;
        }

        // Apply updated mappings
        let updatedCount = 0;
        for (const [key, mapping] of Object.entries(modalRes.mappings)) {
          let keyUpdated = false;
          if (mapping.isFilter) {
            if (mapping.type === 'column') {
              currentSourceConfig.filter_column = mapping.value;
              keyUpdated = true;
            }
          } else {
            for (const prof of mapping.profiles) {
              const targetP = templateProfiles.find(p => p.id === prof.id);
              if (targetP) {
                if (!targetP.field_mappings) targetP.field_mappings = { dynamic: {}, fixed: {} };
                if (!targetP.field_mappings.dynamic) targetP.field_mappings.dynamic = {};
                if (!targetP.field_mappings.fixed) targetP.field_mappings.fixed = {};

                if (mapping.type === 'column') {
                  targetP.field_mappings.dynamic[key] = mapping.value;
                  delete targetP.field_mappings.fixed[key];
                  keyUpdated = true;
                } else if (mapping.type === 'fixed') {
                  targetP.field_mappings.fixed[key] = mapping.value;
                  delete targetP.field_mappings.dynamic[key];
                  keyUpdated = true;
                } else if (mapping.type === 'skip') {
                  delete targetP.field_mappings.dynamic[key];
                  keyUpdated = true;
                }
              }
            }
          }
          if (keyUpdated) {
            updatedCount++;
          }
        }

        if (modalRes.remember) {
          for (const p of templateProfiles) {
            await window.StorageUtils.saveProfile(p);
          }
          const globalConfig = window.AppConfig.get();
          globalConfig.source = currentSourceConfig;
          localStorage.setItem('coupang_config', JSON.stringify(globalConfig));
          localStorage.setItem('my_family_config', JSON.stringify(globalConfig));
          if (typeof renderTemplateProfilesUI === 'function') renderTemplateProfilesUI();
          if (typeof renderSourceConfig === 'function') renderSourceConfig(currentSourceConfig);
        }

        logMessage(`[欄位智慧補全] 已成功補齊/更新 ${updatedCount} 個欄位對應關係！`, 'success');
      }

      sourceExcelFile = file;
      loadedWorkbook = wb;
      excelDropZone.classList.add('has-file');
      excelFileInfo.textContent = `已載入: ${file.name}`;
      excelFileInfo.classList.remove('hidden');
      checkReady();

      if (inputRowEnd) {
        inputRowEnd.placeholder = `最大 ${totalR} 列`;
      }
      updateRangeHintUI();
      const stats = calculateValidProducts(ws);
      const tmplNames = detectedProfiles.map(p => p.name).join('、');
      logMessage(`已載入來源 Excel: ${file.name}（工作表「${ws.name()}」共 ${totalR} 列，自動偵測套用模板：【${tmplNames}】，符合條件之有效商品共 ${stats.count} 筆）`, 'success');

      await autoScanAndSyncMappings(true);
    } catch (err) {
      console.warn('解析 Excel 失敗:', err);
      logMessage(`解析 Excel 失敗: ${err.message}`, 'error');
    }
  }

  // Dynamic Template & Output Processing Pipeline
  btnStartProcess.addEventListener('click', async () => {
    if (!sourceExcelFile || photoFilesArray.length === 0) {
      alert('請先上傳來源 Excel 檔案及選取照片資料夾！');
      return;
    }

    btnStartProcess.disabled = true;
    btnSaveToFolder.disabled = true;
    btnDownloadZip.disabled = true;
    resultSection.classList.remove('hidden');
    progressContainer.classList.remove('hidden');
    setProgress(0, '正在初始化處理引擎與載入模板...');

    processedResults = [];
    generatedExcelFiles = new Map();
    filesToExport = {};

    try {
      const arrayBuffer = await sourceExcelFile.arrayBuffer();
      const wb = await XlsxPopulate.fromDataAsync(arrayBuffer);
      const ws = getSourceSheet(wb);
      if (!ws) throw new Error('來源 Excel 中找不到任何可用的工作表！');

      const headerRow = currentSourceConfig.header_row || 3;
      const rowStartInput = inputRowStart ? inputRowStart.value : null;
      const rowEndInput = inputRowEnd ? inputRowEnd.value : null;
      const rangeInfo = getValidRowsInfo(ws, rowStartInput, rowEndInput);
      const validRowIndices = rangeInfo.validRows;
      const rowStart = rangeInfo.startRow;
      const rowEnd = rangeInfo.endRow;
      const totalValidCount = validRowIndices.length;
      const totalRows = rangeInfo.maxRow;

      const headers = buildHeaderMap(ws, headerRow);
      const nameColIdx = findHeaderColIdx(headers, '商品名稱') || findHeaderColIdx(headers, '中文品名') || findHeaderColIdx(headers, '品名') || findHeaderColIdx(headers, 'NAME');
      const filterColName = currentSourceConfig.filter_column || '中文背標';
      const filterColIdx = findHeaderColIdx(headers, filterColName);

      if (!nameColIdx) {
        throw new Error(`來源 Excel 表頭（第 ${headerRow} 列）缺少「商品名稱」或「中文品名」必要欄位！`);
      }

      logMessage(`處理範圍：第 ${rowStart} 列 至 第 ${rowEnd} 列（共掃描 ${rangeInfo.totalDataRows} 列，其中符合背標條件之有效商品共 ${totalValidCount} 筆，工作表總列數: ${totalRows} 列）`);

      if (totalValidCount === 0) {
        setProgress(100, '所選範圍內無符合條件之有效商品資料。');
        logMessage('未找到任何符合篩選條件的有效商品資料。', 'warning');
        return;
      }

      const processor = new window.MyFamilyProcessor(
        photoFilesArray,
        arrayBuffer,
        currentColorAliases,
        currentCollectionAliases,
        templateProfiles
      );

      // Map to hold loaded Template Workbooks: Map<profileId, { wb, ws, headerMap, nextRowIdx, profile, targetSubfolder, fileName }>
      const activeWorkbooks = new Map();

      async function getOrInitWorkbook(profileId, categoryHint, subfolderHint) {
        if (activeWorkbooks.has(profileId)) {
          return activeWorkbooks.get(profileId);
        }

        let profile = templateProfiles.find(p => p.id === profileId || p.template_type === profileId);
        if (!profile) {
          profile = templateProfiles.find(p => p.id === 'LEASH') || templateProfiles[0];
        }

        let b64 = null;
        if (profile.is_builtin) {
          b64 = (profile.id === 'HARNESS') ? (window.CoupangTemplates?.HARNESS || window.MyFamilyTemplates?.HARNESS) : (window.CoupangTemplates?.LEASH || window.MyFamilyTemplates?.LEASH);
        }
        if (!b64) {
          b64 = await window.StorageUtils.getTemplateData(profile.id);
        }

        let wbInstance = null;
        let wsInstance = null;
        if (b64) {
          wbInstance = await XlsxPopulate.fromDataAsync(b64ToArrayBuffer(b64));
          wsInstance = wbInstance.sheets().find(s => s.name().startsWith('QF_')) || wbInstance.sheet(0);
        } else {
          wbInstance = await XlsxPopulate.fromBlankAsync();
          wsInstance = wbInstance.sheet(0).name(`QF_${profile.name || 'Output'}`);
        }

        const hMap = buildHeaderMap(wsInstance, 5);
        const targetSubfolder = subfolderHint || profile.subfolder || profile.name || '未分類品項';
        const fileName = (profile.template_file_name || `${profile.name}.xlsx`).replace(/\.xlsx$/i, '_auto_generate.xlsx');

        const item = {
          profileId: profile.id,
          profile: profile,
          wb: wbInstance,
          ws: wsInstance,
          headerMap: hMap,
          nextRowIdx: 9,
          targetSubfolder: targetSubfolder,
          fileName: fileName,
          count: 0
        };

        activeWorkbooks.set(profileId, item);
        return item;
      }

      let totalCount = 0;
      let matchCount = 0;
      let missCount = 0;
      const typeCounts = {};

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

        const brandIdx = findHeaderColIdx(headers, 'BRAND') || findHeaderColIdx(headers, '品牌') || findHeaderColIdx(headers, 'Brand');
        const rawBrand = brandIdx ? (getCellValue(ws.cell(r, brandIdx)) || '').toString().trim() : '';

        const collIdx = findHeaderColIdx(headers, 'COLLECTION') || findHeaderColIdx(headers, '系列') || findHeaderColIdx(headers, 'Collection');
        const rawCollection = collIdx ? (getCellValue(ws.cell(r, collIdx)) || '').toString().trim() : '';
        const typeIdx = findHeaderColIdx(headers, 'TYPE') || findHeaderColIdx(headers, '種類') || findHeaderColIdx(headers, '品類') || findHeaderColIdx(headers, 'Type');
        const rawType = typeIdx ? (getCellValue(ws.cell(r, typeIdx)) || '').toString().trim() : '';
        const colorIdx = findHeaderColIdx(headers, 'COLOR') || findHeaderColIdx(headers, '顏色') || findHeaderColIdx(headers, 'Color');
        const rawColor = colorIdx ? (getCellValue(ws.cell(r, colorIdx)) || '').toString().trim() : '';
        const sizeIdx = findHeaderColIdx(headers, 'SIZE') || findHeaderColIdx(headers, '尺寸') || findHeaderColIdx(headers, 'Size');
        const rawSize = sizeIdx ? (getCellValue(ws.cell(r, sizeIdx)) || '').toString().trim() : '';
        
        const rawHints = { brand: rawBrand, collection: rawCollection, type: rawType, color: rawColor, size: rawSize, sku: skuStr };
        const parsed = processor.parseChineseName(zhName, rawSize, rawHints);
        const targetInfo = processor.getTargetTemplateAndCategory(zhName, parsed.type || rawType);
        totalCount++;

        let isUnmatched = targetInfo.unmatched;
        if (isUnmatched) {
          logMessage(`[警告] 商品「${zhName}」未命中任何品類規則，尚未加入該品項的模板！將標記紅底警示`, 'warning');
        }

        let imgs = processor.getLocalImagesForProduct(parsed, rawHints);
        if (!parsed.success) {
          logMessage(`[提示] 商品「${zhName}」採彈性回退解析 (結合 Excel 欄位資訊)`, 'info');
        }

        const wbItem = await getOrInitWorkbook(targetInfo.template_id, targetInfo.category, targetInfo.target_subfolder);
        wbItem.count++;
        typeCounts[wbItem.profile.name] = (typeCounts[wbItem.profile.name] || 0) + 1;

        const targetSubFolder = wbItem.targetSubfolder;
        if (!filesToExport[targetSubFolder]) {
          filesToExport[targetSubFolder] = {
            images: new Map(),
            back_labels: new Map(),
            excelFiles: new Map()
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
              const blob = await window.ImageUtils.ensureMinShortEdge(await window.ImageUtils.loadImage(imgs.chart));
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
            const blob = await window.ImageUtils.ensureMinShortEdge(await window.ImageUtils.loadImage(imgs.label));
            const szSuffix = (rawSize || parsed.size) ? `_${String(rawSize || parsed.size).trim()}` : '';
            const cleanLabelName = sanitizeFilename(`背標_${zhName}${szSuffix}`);
            labelImgName = `${cleanLabelName}.jpg`;
            filesToExport[targetSubFolder].back_labels.set(labelImgName, blob);
          } catch(imgErr) {
            logMessage(`[警告] 背標處理失敗: ${imgs.label.name}`, 'warning');
          }
        }

        if (mainImgName || imgs.main) {
          matchCount++;
        } else {
          missCount++;
          logMessage(`[警告] 找不到主要圖片: SKU=${skuStr}, 品名=${zhName}`, 'warning');
        }

        const targetWs = wbItem.ws;
        const targetRowIdx = wbItem.nextRowIdx++;
        const tmplMap = wbItem.headerMap;
        const profileMappings = wbItem.profile.field_mappings || {};

        const getSourceHeaderIdx = (colName) => findHeaderColIdx(headers, colName);

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

            if (isUnmatched || !parsed.success || ((!mainImgName && !imgs.main) && ['商品名稱', '商品正面(主要圖片）'].includes(colName))) {
              cell.style('fill', 'ffff0000');
            }
          }
        };

        // Write category
        const catValue = isUnmatched ? '' : (targetInfo.category || wbItem.profile.category_name || '');
        setVal('細分商品種類', catValue);

        // Write Fixed mappings from profile
        for (const [tKey, fVal] of Object.entries(profileMappings.fixed || {})) {
          if (['細分商品種類'].includes(tKey)) continue;
          setVal(tKey, fVal);
        }

        // Write Dynamic mappings from profile
        for (const [tKey, sCol] of Object.entries(profileMappings.dynamic || {})) {
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

        // Auto parsed attributes with generic fallback
        const finalBrand = parsed.brand || rawBrand;
        const finalCollection = parsed.collection || rawCollection;
        const finalColor = parsed.color || rawColor || processor.extractColorFromZhName(zhName, rawSize);
        const finalSize = parsed.size || rawSize;

        if (finalBrand) {
          setVal('品牌', finalBrand);
          if (!profileMappings.fixed?.['製造廠商']) {
            setVal('製造廠商', finalBrand);
          }
        }
        if (finalCollection) setVal('系列', finalCollection);
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

      // Generate binary buffers for all active workbooks
      setProgress(95, '正在壓縮並修正 Excel 格式...');
      for (const [profId, item] of activeWorkbooks.entries()) {
        if (item.count > 0) {
          let buf = await item.wb.outputAsync();
          buf = await patchXlsxDimension(buf);
          if (!filesToExport[item.targetSubfolder]) {
            filesToExport[item.targetSubfolder] = { images: new Map(), back_labels: new Map(), excelFiles: new Map() };
          }
          filesToExport[item.targetSubfolder].excelFiles.set(item.fileName, buf);
        }
      }

      setProgress(100, '處理完成！請選擇匯出方式。');
      statTotalSku.textContent = totalCount;
      statMatchedImages.textContent = matchCount;
      statMissingImages.textContent = missCount;

      if (totalCount === 0) {
        btnSaveToFolder.disabled = true;
        btnDownloadZip.disabled = true;
        logMessage('處理完成，但未找到任何符合篩選條件的商品資料。', 'warning');
      } else {
        btnSaveToFolder.disabled = false;
        btnDownloadZip.disabled = false;
        const typesDesc = [];
        for (const [name, count] of Object.entries(typeCounts)) {
          typesDesc.push(`【${name}】: ${count} 筆`);
        }
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

  // Save to Folder (File System Access API)
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
      let totalWriteCount = 0;

      for (const [folderName, folderData] of Object.entries(filesToExport)) {
        const hasContent = folderData.images.size > 0 || folderData.back_labels.size > 0 || folderData.excelFiles.size > 0;
        if (!hasContent) continue;
        totalWriteCount += folderData.excelFiles.size + folderData.images.size + folderData.back_labels.size;
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

      let totalFiles = 0;
      for (const [folderName, folderData] of Object.entries(filesToExport)) {
        const hasContent = folderData.images.size > 0 || folderData.back_labels.size > 0 || folderData.excelFiles.size > 0;
        if (!hasContent) continue;

        savedCategories.push(folderName);
        const subDir = await dirHandle.getDirectoryHandle(folderName, { create: true });
        
        for (const [fname, buf] of folderData.excelFiles.entries()) {
          const fh = await subDir.getFileHandle(fname, { create: true });
          const w = await fh.createWritable();
          await w.write(buf);
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
      logMessage(`全部處理完成！已儲存【${savedCategories.join('、')}】資料，共 ${totalFiles} 個不重複圖檔，無安全性警告！`, 'success');
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

  // Download ZIP
  btnDownloadZip.addEventListener('click', async () => {
    if (processedResults.length === 0) return;

    progressContainer.classList.remove('hidden');
    setProgress(0, '正在準備 ZIP 檔案結構...');

    const zip = new JSZip();
    const exportedCategories = [];

    for (const [folderName, folderData] of Object.entries(filesToExport)) {
      const hasContent = folderData.images.size > 0 || folderData.back_labels.size > 0 || folderData.excelFiles.size > 0;
      if (!hasContent) continue;

      exportedCategories.push(folderName);
      const root = zip.folder(folderName);
      
      for (const [fname, buf] of folderData.excelFiles.entries()) {
        root.file(fname, buf);
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

  // Download sample structure
  btnDownloadSample.addEventListener('click', async () => {
    const zip = new JSZip();
    const photoFolder = zip.folder("Photo");
    const hermitage = photoFolder.folder("HERMITAGE");
    const dogCollar = hermitage.folder("DOG COLLAR");
    const blackCollar = dogCollar.folder("BLACK");
    blackCollar.file("HERMITAGE_DOG_COLLAR_BLACK_主圖.jpg", "（請放入 1000x1000 商品正面主圖）");
    blackCollar.file("1.jpg", "（請放入情境圖 1）");
    blackCollar.file("2.jpg", "（請放入情境圖 2）");
    dogCollar.file("尺寸規格表_HERMITAGE_DOG_COLLAR.jpg", "（請放入尺寸規格圖）");
    const labels = photoFolder.folder("背標");
    labels.file("背標_品牌名 隱士系列 頂級皮革狗項圈 經典黑 L_L.jpg", "（請放入商品背標照片）");

    const content = await zip.generateAsync({ type: "blob" });
    saveAs(content, "Coupang_Photo_Sample_Structure.zip");
    logMessage("範例資料夾結構已下載！", "success");
  });
});
