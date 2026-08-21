const APP_VERSION = 'v2.16.1';

function normalizeHeaderKey(str) {
  if (window.SharedUtils) return window.SharedUtils.normalizeKey(str);
  if (!str) return '';
  return String(str)
    .normalize('NFKC')
    .replace(/[\uFEFF\u200B]/g, '')
    .replace(/[\u3000\u00A0]/g, ' ')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .trim();
}

/**
 * Helper to escape HTML characters
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Extract all non-empty column header names from a given worksheet at headerRow
 */
function getSheetHeadersList(ws, headerRow = 3) {
  if (!ws) return [];
  const headers = [];
  const range = ws.usedRange();
  const maxCol = range ? range.endCell().columnNumber() : 70;
  for (let c = 1; c <= maxCol; c++) {
    const val = ws.cell(headerRow, c).value();
    if (val !== null && val !== undefined) {
      const str = String(val).trim();
      if (str && !headers.includes(str)) {
        headers.push(str);
      }
    }
  }
  return headers;
}

/**
 * Standard candidate source Excel header names
 */
function getCommonSourceHeaders() {
  return [
    '中文背標', '中文品名', '商品名稱', '產品中文名稱', 'TYPE', '品類', '種類', '款式',
    'COLLECTION', '系列', '中文顏色', 'COLOR', '顏色', 'SIZE', '尺寸',
    '國際條碼', '原廠貨號', '定價', '售價', '建議售價', '產地', '材質',
    '適用對象', '洗滌方式', '規格', '重量', '淨重', '條碼'
  ];
}

/**
 * Get all available source header names from active workbook sheet or standard fallbacks
 */
function getAvailableSourceHeaders(preferredWs = null, preferredHeaderRow = 3) {
  if (preferredWs) {
    const h = getSheetHeadersList(preferredWs, preferredHeaderRow);
    if (h.length > 0) return h;
  }
  if (typeof loadedWorkbook !== 'undefined' && loadedWorkbook) {
    const ws = (typeof currentSourceConfig !== 'undefined' && currentSourceConfig?.sheet_name)
      ? (loadedWorkbook.sheet(currentSourceConfig.sheet_name) || loadedWorkbook.sheet(0))
      : loadedWorkbook.sheet(0);
    if (ws) {
      const h = getSheetHeadersList(ws, (typeof currentSourceConfig !== 'undefined' && currentSourceConfig?.header_row) || 3);
      if (h.length > 0) return h;
    }
  }
  return getCommonSourceHeaders();
}

/**
 * Build HTML option tags for column select dropdowns
 */
function buildColSelectOptions(headers, currentVal, emptyLabel = '-- 請選擇來源欄位 (未指定則留空) --') {
  let html = `<option value="">${emptyLabel}</option>`;
  const seen = new Set();
  const validHeaders = Array.isArray(headers) && headers.length > 0 ? headers : getCommonSourceHeaders();
  
  if (currentVal && !validHeaders.some(h => h.toUpperCase() === String(currentVal).toUpperCase())) {
    html += `<option value="${escapeHtml(currentVal)}" selected>${escapeHtml(currentVal)} (自訂)</option>`;
    seen.add(currentVal);
  }

  validHeaders.forEach(h => {
    seen.add(h);
    const isSel = (currentVal && h.toUpperCase() === String(currentVal).toUpperCase()) ? 'selected' : '';
    html += `<option value="${escapeHtml(h)}" ${isSel}>${escapeHtml(h)}</option>`;
  });
  return html;
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
 * Scan target worksheet and highlight empty cells in '必填' (Required) columns with solid red fill
 */
function highlightMissingRequiredCells(worksheet, startRow = 9, endRow = null, reqRow = 6) {
  if (!worksheet) return;
  const maxCol = worksheet.usedRange() ? worksheet.usedRange().endCell().columnNumber() : 70;
  const lastRow = (endRow !== null && endRow !== undefined) ? endRow : (worksheet.usedRange() ? worksheet.usedRange().endCell().rowNumber() : startRow);

  // Find all column indices that are marked as '必填' in reqRow (Row 6)
  const requiredColIndices = [];
  for (let c = 1; c <= maxCol; c++) {
    const reqVal = worksheet.cell(reqRow, c).value();
    if (reqVal && String(reqVal).trim().includes('必填')) {
      requiredColIndices.push(c);
    }
  }

  if (requiredColIndices.length === 0) return;

  // Scan all data rows and highlight empty required cells
  for (let r = startRow; r < lastRow; r++) {
    for (const colIdx of requiredColIndices) {
      const cell = worksheet.cell(r, colIdx);
      const val = cell.value();
      const isEmpty = (val === undefined || val === null || (typeof val === 'string' && val.trim() === '') || (typeof val === 'number' && isNaN(val)));
      if (isEmpty) {
        try {
          cell.style('fill', 'ffff0000');
        } catch(e) {
          console.warn(`[Highlight] Failed to style cell at (${r}, ${colIdx}):`, e);
        }
      }
    }
  }
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
        // Check global fixed in source config
        const globalFixed = (typeof currentSourceConfig !== 'undefined' && currentSourceConfig?.fixed) ? currentSourceConfig.fixed : (window.AppConfig?.getDefaultSourceConfig().fixed || {});
        let foundGlobalFix = null;
        for (const [k, v] of Object.entries(globalFixed)) {
          if (normalizeHeaderKey(k) === normalizeHeaderKey(colName) || k.replace(/\s+/g, '') === colName.replace(/\s+/g, '')) {
            foundGlobalFix = v;
            break;
          }
        }

        if (foundGlobalFix !== null) {
          status = 'inherited';
          mappingType = 'fixed';
          defaultValue = foundGlobalFix;
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

function sanitizeFilename(name, fallbackName = '未命名檔案') {
  if (window.SharedUtils) return window.SharedUtils.sanitizeFilename(name, fallbackName);
  if (!name) return fallbackName;
  let s = String(name)
    .normalize('NFKC')
    .replace(/[\\/:*?"<>|\r\n\t\0]/g, '_')
    .replace(/[.\s]+$/, '')
    .trim();
  return s || fallbackName;
}

function formatBarcode(val) {
  if (window.SharedUtils) return window.SharedUtils.formatBarcode(val);
  if (val === null || val === undefined || val === '') return '';
  let s = String(val).normalize('NFKC').trim();
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
  if (window.SharedUtils) return window.SharedUtils.getCellValue(cell);
  if (!cell) return '';
  const val = cell.value();
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && val.text) return val.text;
  return val;
}

function getBaseProductName(fullName, size) {
  if (!fullName) return '';
  let s = (window.SharedUtils ? window.SharedUtils.normalizeKey(fullName) : String(fullName).normalize('NFKC').trim());
  if (size) {
    const szStr = (window.SharedUtils ? window.SharedUtils.normalizeKey(size) : String(size).normalize('NFKC').trim());
    if (szStr && s.endsWith(szStr)) {
      s = s.slice(0, -szStr.length).trim();
    }
  }
  return s;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Scan data rows in source Excel to dynamically detect which template profiles are actively needed
 * and return unmatched products if any.
 */
function detectRequiredTemplates(ws, headerRow, rowStart, maxRow, profiles, sourceConfig = null) {
  if (!ws || !profiles || profiles.length === 0) {
    return { detectedProfiles: profiles || [], unmatchedItems: [] };
  }
  const headerMap = buildHeaderMap(ws, headerRow);
  let nameColIdx = null;
  for (const p of profiles) {
    const dynName = p.field_mappings?.dynamic?.['商品名稱'] || p.field_mappings?.dynamic?.['中文品名'] || p.field_mappings?.dynamic?.['產品中文名稱'];
    if (dynName) {
      nameColIdx = findHeaderColIdx(headerMap, dynName);
      if (nameColIdx) break;
    }
  }
  if (!nameColIdx) {
    const nameCandidates = ['商品名稱', '中文品名', '產品中文名稱', '產品名稱', '商品中文名稱', '中文名稱', '品名(中)', '品名', '商品名', '商品', 'NAME', 'Item Name', 'Product Name', 'Description'];
    for (const cand of nameCandidates) {
      nameColIdx = findHeaderColIdx(headerMap, cand);
      if (nameColIdx) break;
    }
  }
  const typeColIdx = (sourceConfig?.type_column ? findHeaderColIdx(headerMap, sourceConfig.type_column) : null) || findHeaderColIdx(headerMap, 'TYPE') || findHeaderColIdx(headerMap, '種類') || findHeaderColIdx(headerMap, '品類') || findHeaderColIdx(headerMap, 'Type');
  const filterColName = (sourceConfig?.filter_column !== undefined) ? sourceConfig.filter_column : '';
  const filterColIdx = (filterColName && filterColName.trim() !== '') ? findHeaderColIdx(headerMap, filterColName) : null;

  const matchedProfileIds = new Set();
  const unmatchedItems = [];
  const processor = new window.CoupangProcessor([], null, {}, {}, profiles);

  const range = ws.usedRange();
  const totalR = range ? range.endCell().rowNumber() : 100;
  const actualStart = Math.max(headerRow + 1, rowStart || (headerRow + 1));
  const actualEnd = maxRow ? Math.min(maxRow, totalR) : totalR;

  // Fallback: discover name column by scanning cell contents against known category keywords
  if (!nameColIdx) {
    const allKeywords = (typeof CoupangProcessor !== 'undefined' && CoupangProcessor.getAllCategoryKeywords)
      ? CoupangProcessor.getAllCategoryKeywords(profiles)
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
    if (targetInfo && targetInfo.template_id && targetInfo.template_id !== 'UNMATCHED' && !targetInfo.unmatched) {
      matchedProfileIds.add(targetInfo.template_id);
    } else {
      unmatchedItems.push({
        row: r,
        name: zhName,
        type: rawType
      });
    }
  }

  const detected = profiles.filter(p => matchedProfileIds.has(p.id) || matchedProfileIds.has(p.template_type));
  return {
    detectedProfiles: detected.length > 0 ? detected : profiles,
    unmatchedItems
  };
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
    { keywords: ['品名', '商品名稱', '中文品名', '產品名稱', '產品中文名稱', '商品中文名稱', '中文名稱', '品名(中)', 'NAME', 'ITEMNAME', 'DESCRIPTION', '商品名', '中文', '商品'], targets: ['商品名稱', '中文品名', '品名', '產品名稱', '產品中文名稱'] },
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
  const filterCol = (sourceConfig?.filter_column !== undefined) ? sourceConfig.filter_column : '';
  if (filterCol && filterCol.trim() !== '') {
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
  let nameIdx = null;
  for (const p of detectedProfiles) {
    const dynName = p.field_mappings?.dynamic?.['商品名稱'] || p.field_mappings?.dynamic?.['中文品名'] || p.field_mappings?.dynamic?.['產品中文名稱'];
    if (dynName) {
      nameIdx = findHeaderColIdx(headerMap, dynName);
      if (nameIdx) break;
    }
  }
  if (!nameIdx) {
    const nameCandidates = ['商品名稱', '中文品名', '產品中文名稱', '產品名稱', '商品中文名稱', '中文名稱', '品名(中)', '品名', '商品名', '商品', 'NAME', 'Item Name', 'Product Name'];
    for (const cand of nameCandidates) {
      nameIdx = findHeaderColIdx(headerMap, cand);
      if (nameIdx) break;
    }
  }
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
  const statMissingBackLabels = document.getElementById('statMissingBackLabels');
  const statUnmatchedTemplates = document.getElementById('statUnmatchedTemplates');
  
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
  const wizardScannedCount = document.getElementById('wizardScannedCount');
  const wizardTableBody = document.getElementById('wizardTableBody');

  // Sidebar Unmatched Type Hint elements
  const unmatchedTypeHintContainer = document.getElementById('unmatchedTypeHintContainer');
  const unmatchedTypePillsList = document.getElementById('unmatchedTypePillsList');

  // Excel Upload Wizard elements (兩步驟載入精靈)
  const excelUploadWizardModal = document.getElementById('excelUploadWizardModal');
  const uploadWizardClose = document.getElementById('uploadWizardClose');
  const uploadWizardTitle = document.getElementById('uploadWizardTitle');
  const wizardStepIndicator1 = document.getElementById('wizardStepIndicator1');
  const wizardStepIndicator2 = document.getElementById('wizardStepIndicator2');
  const wizardStepperLine = document.getElementById('wizardStepperLine');
  const uploadWizardStep1 = document.getElementById('uploadWizardStep1');
  const uploadWizardStep2 = document.getElementById('uploadWizardStep2');
  const uploadWizardFileName = document.getElementById('uploadWizardFileName');
  const wizardSheetSelect = document.getElementById('wizardSheetSelect');
  const wizardHeaderRow = document.getElementById('wizardHeaderRow');
  const wizardRowStart = document.getElementById('wizardRowStart');
  const wizardFilterColumn = document.getElementById('wizardFilterColumn');
  const detectedTemplatesContainer = document.getElementById('detectedTemplatesContainer');
  const wizardUnmatchedAlertBox = document.getElementById('wizardUnmatchedAlertBox');
  const wizardUnmatchedTypeCount = document.getElementById('wizardUnmatchedTypeCount');
  const wizardUnmatchedTotalCount = document.getElementById('wizardUnmatchedTotalCount');
  const wizardUnmatchedAccordionList = document.getElementById('wizardUnmatchedAccordionList');
  const btnWizardGoToConfig = document.getElementById('btnWizardGoToConfig');
  const wizardBrandFixed = document.getElementById('wizardBrandFixed');
  const wizardManufacturerFixed = document.getElementById('wizardManufacturerFixed');
  const wizardCollectionColumn = document.getElementById('wizardCollectionColumn');
  const wizardTypeColumn = document.getElementById('wizardTypeColumn');
  const wizardColorColumn = document.getElementById('wizardColorColumn');
  const wizardSizeColumn = document.getElementById('wizardSizeColumn');
  const missingColumnCount = document.getElementById('missingColumnCount');
  const missingColumnsContainer = document.getElementById('missingColumnsContainer');
  const mappingTableBody = document.getElementById('mappingTableBody');
  const allColumnsMatchedNotice = document.getElementById('allColumnsMatchedNotice');
  const chkRememberMappings = document.getElementById('chkRememberMappings');
  const btnWizardPrev = document.getElementById('btnWizardPrev');
  const btnWizardCancel = document.getElementById('btnWizardCancel');
  const btnWizardNext = document.getElementById('btnWizardNext');
  const btnWizardConfirm = document.getElementById('btnWizardConfirm');

  // State Variables
  let templateProfiles = [];
  let activeProfileId = 'HARNESS';
  let pendingWizard = null;
  let lastDetectedUnmatchedTypes = [];

  let currentSourceConfig = window.AppConfig.get().source;

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

        // Auto-migrate stored profiles to ensure baseline fixed keys exist
        for (const p of templateProfiles) {
          if (!p.field_mappings) p.field_mappings = { fixed: {}, dynamic: {} };
          if (!p.field_mappings.fixed) p.field_mappings.fixed = {};
          let modified = false;
          if (!p.field_mappings.fixed['法定種類'] || p.field_mappings.fixed['法定種類'] === '') {
            p.field_mappings.fixed['法定種類'] = 'TW_General';
            modified = true;
          }
          if (modified) {
            await window.StorageUtils.saveProfile(p);
          }
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

    // 渲染設定檔清單下方缺少 Type 之紅框橢圓形標籤提示
    if (unmatchedTypeHintContainer && unmatchedTypePillsList) {
      if (Array.isArray(lastDetectedUnmatchedTypes) && lastDetectedUnmatchedTypes.length > 0) {
        unmatchedTypeHintContainer.classList.remove('hidden');
        unmatchedTypePillsList.innerHTML = '';
        lastDetectedUnmatchedTypes.forEach(item => {
          const pill = document.createElement('div');
          pill.className = 'unmatched-type-pill';
          pill.title = `來源 Excel 中有 ${item.count} 筆商品屬於「${item.type}」，尚未設定專屬模板`;
          pill.innerHTML = `
            <span class="material-symbols-outlined" style="font-size: 0.85rem; color: #dc2626;">label</span>
            <span>${escapeHtml(item.type)}</span>
            <span class="pill-count">(${item.count}筆)</span>
          `;
          unmatchedTypePillsList.appendChild(pill);
        });
      } else {
        unmatchedTypeHintContainer.classList.add('hidden');
      }
    }

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

    const availableHeaders = getAvailableSourceHeaders();

    let fixedHtml = '';
    for (const [k, v] of Object.entries(profile.field_mappings?.fixed || {})) {
      fixedHtml += `
        <div class="dynamic-row fixed-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${escapeHtml(k)}">
          <input type="text" class="val-input" placeholder="固定填寫內容" value="${escapeHtml(v)}">
          <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-symbols-outlined">delete</span></button>
        </div>`;
    }

    let dynamicHtml = '';
    for (const [k, v] of Object.entries(profile.field_mappings?.dynamic || {})) {
      dynamicHtml += `
        <div class="dynamic-row dynamic-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${escapeHtml(k)}">
          <select class="val-input form-select">
            ${buildColSelectOptions(availableHeaders, v, '-- 請選擇來源表對應欄位 --')}
          </select>
          <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-symbols-outlined">delete</span></button>
        </div>`;
    }

    let currentProfileKeywords = Array.isArray(profile.keywords)
      ? [...profile.keywords]
      : (profile.keywords ? String(profile.keywords).split(/[,，]/).map(k => k.trim()).filter(Boolean) : (profile.name ? [profile.name] : []));

    profileEditorContent.innerHTML = `
      <div class="profile-meta-card">
        <div class="profile-meta-grid">
          <div class="input-row" style="margin:0;">
            <label style="flex:0 0 110px; font-weight:600;">設定檔名稱:</label>
            <input type="text" id="prof_name" value="${escapeHtml(profile.name || '')}" ${profile.is_builtin ? 'readonly style="background:#f1f5f9;"' : ''}>
          </div>
          <div class="input-row" style="margin:0; align-items:flex-start;">
            <label style="flex:0 0 110px; font-weight:600; padding-top:6px;">匹配關鍵字:</label>
            <div class="keyword-tags-wrapper">
              <div class="keyword-input-group">
                <input type="text" id="profKeywordInput" placeholder="輸入關鍵字後點擊 + 或按 Enter 新增 (支援逗號分隔)...">
                <button type="button" id="btnAddProfKeyword" class="btn btn-secondary btn-sm btn-icon" title="新增關鍵字"><span class="material-symbols-outlined">add</span></button>
              </div>
              <div id="profKeywordsTagsContainer" class="keyword-tags-container"></div>
            </div>
          </div>
          <div class="input-row" style="margin:0;">
            <label style="flex:0 0 110px; font-weight:600;">預設分類代碼:</label>
            <input type="text" id="prof_category_name" placeholder="例如: 寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)" value="${escapeHtml(profile.category_name || '')}">
          </div>
        </div>
        <div class="template-file-bar" style="margin-top:12px;">
          <div class="template-file-info">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
            <span>Excel 範本: <strong>${escapeHtml(profile.template_file_name || (profile.id + '.xlsx'))}</strong></span>
            <span class="badge-${profile.is_builtin ? 'system' : 'inherited'}">${profile.is_builtin ? '內建範本' : '自訂上傳'}</span>
          </div>
          <div class="template-file-actions">
            <button type="button" id="btnDownloadProfileTemplate" class="btn btn-outline btn-sm btn-icon" title="下載 Excel 範本檔"><span class="material-symbols-outlined">download</span></button>
            <button type="button" id="btnReplaceProfileTemplate" class="btn btn-secondary btn-sm btn-icon" title="替換 Excel 範本檔"><span class="material-symbols-outlined">upload</span></button>
            <input type="file" id="replaceTemplateFileInput" accept=".xlsx" hidden>
          </div>
        </div>
      </div>

      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 1px solid #e2e8f0; margin-bottom:10px; padding-bottom:6px;">
          <h4 style="border:none; margin:0; padding:0;">固定欄位對應 (Fixed Mappings)</h4>
          <button type="button" class="btn btn-outline btn-sm btn-icon" title="新增固定欄位" onclick="addProfileFixedRow()"><span class="material-symbols-outlined">add</span></button>
        </div>
        <p style="font-size:0.82rem; color:#64748b; margin-bottom:8px;">不論來源資料為何，強制填入目標 Excel 模板的固定值。</p>
        <div id="profileFixedContainer">${fixedHtml}</div>
      </div>

      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 1px solid #e2e8f0; margin-bottom:10px; padding-bottom:6px;">
          <h4 style="border:none; margin:0; padding:0;">動態欄位對應 (Dynamic Mappings)</h4>
          <button type="button" class="btn btn-outline btn-sm btn-icon" title="新增動態欄位" onclick="addProfileDynamicRow()"><span class="material-symbols-outlined">add</span></button>
        </div>
        <p style="font-size:0.82rem; color:#64748b; margin-bottom:8px;">將來源商品表的欄位資料，動態填入目標 Excel 模板的對應欄位中。</p>
        <div id="profileDynamicContainer">${dynamicHtml}</div>
      </div>

      ${!profile.is_builtin ? `
        <div style="display:flex; justify-content:flex-end; margin-top:20px; padding-top:10px; border-top:1px dashed #e2e8f0;">
          <button type="button" id="btnDeleteCurrentProfile" class="btn btn-danger btn-sm btn-icon" title="刪除此模板設定檔"><span class="material-symbols-outlined">delete_forever</span></button>
        </div>
      ` : ''}
    `;

    // Render keyword tag pills
    function renderProfKeywordTags() {
      const container = document.getElementById('profKeywordsTagsContainer');
      if (!container) return;
      container.innerHTML = '';
      if (currentProfileKeywords.length === 0) {
        container.innerHTML = '<span class="keyword-tag-empty">尚未設定關鍵字（輸入關鍵字後點擊 + 或按 Enter 新增）</span>';
        return;
      }
      currentProfileKeywords.forEach((kw, index) => {
        const pill = document.createElement('div');
        pill.className = 'keyword-tag-pill';
        pill.innerHTML = `
          <span class="material-symbols-outlined">label</span>
          <span>${escapeHtml(kw)}</span>
          <button type="button" class="keyword-tag-remove" title="移除關鍵字" data-index="${index}">
            <span class="material-symbols-outlined">close</span>
          </button>
        `;
        pill.querySelector('.keyword-tag-remove').addEventListener('click', (e) => {
          e.stopPropagation();
          currentProfileKeywords.splice(index, 1);
          renderProfKeywordTags();
        });
        container.appendChild(pill);
      });
    }

    const btnAddProfKeyword = document.getElementById('btnAddProfKeyword');
    const profKeywordInput = document.getElementById('profKeywordInput');

    function addProfKeywordFromInput() {
      if (!profKeywordInput) return;
      const val = profKeywordInput.value.trim();
      if (!val) return;
      const items = val.split(/[,，\n]/).map(s => s.trim()).filter(Boolean);
      items.forEach(item => {
        if (!currentProfileKeywords.includes(item)) {
          currentProfileKeywords.push(item);
        }
      });
      profKeywordInput.value = '';
      renderProfKeywordTags();
    }

    if (btnAddProfKeyword) btnAddProfKeyword.addEventListener('click', addProfKeywordFromInput);
    if (profKeywordInput) {
      profKeywordInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          addProfKeywordFromInput();
        }
      });
    }

    renderProfKeywordTags();

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
        if (!b64 && profile.template_type) {
          b64 = await window.StorageUtils.getTemplateData(profile.template_type);
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

    // Attach profile keyword reference to profile for saving
    profile._activeKeywords = currentProfileKeywords;
  }

  window.addProfileFixedRow = function(k='', v='') {
    const div = document.createElement('div');
    div.className = 'dynamic-row fixed-field-row';
    div.innerHTML = `
      <input type="text" class="key-input" placeholder="目標模板欄位" value="${escapeHtml(k)}">
      <input type="text" class="val-input" placeholder="固定填寫內容" value="${escapeHtml(v)}">
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-symbols-outlined">delete</span></button>
    `;
    document.getElementById('profileFixedContainer')?.appendChild(div);
  };

  window.addProfileDynamicRow = function(k='', v='') {
    const availableHeaders = getAvailableSourceHeaders();
    const div = document.createElement('div');
    div.className = 'dynamic-row dynamic-field-row';
    div.innerHTML = `
      <input type="text" class="key-input" placeholder="目標模板欄位" value="${escapeHtml(k)}">
      <select class="val-input form-select">
        ${buildColSelectOptions(availableHeaders, v, '-- 請選擇來源表對應欄位 --')}
      </select>
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-symbols-outlined">delete</span></button>
    `;
    document.getElementById('profileDynamicContainer')?.appendChild(div);
  };

  function saveActiveProfileFromUI() {
    const profile = templateProfiles.find(p => p.id === activeProfileId);
    if (!profile) return;

    const nameInput = document.getElementById('prof_name');
    const categoryInput = document.getElementById('prof_category_name');

    if (nameInput && !profile.is_builtin) profile.name = nameInput.value.trim();
    if (categoryInput) profile.category_name = categoryInput.value.trim();

    if (Array.isArray(profile._activeKeywords)) {
      profile.keywords = profile._activeKeywords.length > 0 ? [...profile._activeKeywords] : [profile.name || ''];
    }

    const fixed = {};
    document.querySelectorAll('#profileFixedContainer .fixed-field-row').forEach(row => {
      const k = row.querySelector('.key-input')?.value.trim();
      const v = row.querySelector('.val-input')?.value.trim();
      if (k) fixed[k] = v;
    });

    const dynamic = {};
    document.querySelectorAll('#profileDynamicContainer .dynamic-field-row').forEach(row => {
      const k = row.querySelector('.key-input')?.value.trim();
      const v = row.querySelector('.val-input')?.value.trim();
      if (k) dynamic[k] = v;
    });

    profile.field_mappings = { fixed, dynamic };
  }

  window.addSourceFixedRow = function(k='', v='') {
    const div = document.createElement('div');
    div.className = 'dynamic-row fixed-field-row';
    div.innerHTML = `
      <input type="text" class="key-input" placeholder="目標模板欄位" value="${escapeHtml(k)}">
      <input type="text" class="val-input" placeholder="全域固定填寫內容" value="${escapeHtml(v)}">
      <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-symbols-outlined">delete</span></button>
    `;
    document.getElementById('sourceFixedContainer')?.appendChild(div);
  };

  // Render Source Config
  function renderSourceConfig(source) {
    if (!guiSourceConfigForm) return;

    let sourceFixedHtml = '';
    const defaultFixed = window.AppConfig?.getDefaultSourceConfig().fixed || {};
    const effectiveFixed = (source.fixed && Object.keys(source.fixed).length > 0) ? source.fixed : defaultFixed;
    for (const [k, v] of Object.entries(effectiveFixed)) {
      sourceFixedHtml += `
        <div class="dynamic-row fixed-field-row">
          <input type="text" class="key-input" placeholder="目標模板欄位" value="${escapeHtml(k)}">
          <input type="text" class="val-input" placeholder="全域固定填寫內容" value="${escapeHtml(v)}">
          <button type="button" class="btn btn-danger btn-sm btn-icon" title="刪除此欄位對應" onclick="this.parentElement.remove()"><span class="material-symbols-outlined">delete</span></button>
        </div>`;
    }

    const allSheets = (loadedWorkbook && loadedWorkbook.sheets) ? loadedWorkbook.sheets() : [];
    let sheetOptionsHtml = '';
    if (allSheets.length > 0) {
      allSheets.forEach(s => {
        const rCount = s.usedRange() ? s.usedRange().endCell().rowNumber() : 0;
        const isSel = s.name() === (source.sheet_name || 'Sheet1') ? 'selected' : '';
        sheetOptionsHtml += `<option value="${escapeHtml(s.name())}" ${isSel}>${escapeHtml(s.name())} (${rCount > 0 ? `共 ${rCount} 列` : '空白工作表'})</option>`;
      });
    } else {
      const defaultSheets = [source.sheet_name || 'Sheet1', 'Sheet1', '商品資料', '工作表1'];
      const uniqueSheets = Array.from(new Set(defaultSheets.filter(Boolean)));
      uniqueSheets.forEach(name => {
        const isSel = name === (source.sheet_name || 'Sheet1') ? 'selected' : '';
        sheetOptionsHtml += `<option value="${escapeHtml(name)}" ${isSel}>${escapeHtml(name)}</option>`;
      });
    }

    let currentSheetHeaders = [];
    if (loadedWorkbook) {
      const ws = loadedWorkbook.sheet(source.sheet_name) || loadedWorkbook.sheet(0);
      if (ws) {
        currentSheetHeaders = getSheetHeadersList(ws, source.header_row || 3);
      }
    }
    if (currentSheetHeaders.length === 0) {
      currentSheetHeaders = getCommonSourceHeaders();
    }

    guiSourceConfigForm.innerHTML = `
      <div class="form-group">
        <h4>來源工作表與行列設定 (Source Sheet & Rows)</h4>
        <div class="input-row">
          <label>來源表名稱:</label>
          <select id="cfg_sheet_name" class="form-select">${sheetOptionsHtml}</select>
        </div>
        <div class="input-row">
          <label>標題列位於第幾列:</label>
          <input type="number" id="cfg_header_row" value="${source.header_row || 3}" min="1">
        </div>
        <div class="input-row">
          <label>資料起始列:</label>
          <input type="number" id="cfg_row_start" value="${source.row_start || 4}" min="1">
        </div>
        <div class="input-row">
          <label>篩選欄位(有填寫才處理):</label>
          <select id="cfg_filter_column" class="form-select">${buildColSelectOptions(currentSheetHeaders, source.filter_column || '', '-- 不啟用篩選 (處理全部列) --')}</select>
        </div>
      </div>

      <div class="form-group">
        <h4>全域屬性與動態對應設定 (Global Attributes & Dynamic Mappings)</h4>
        <p style="font-size:0.82rem; color:#64748b; margin-bottom:12px;">設定整份來源 Excel 通用的固定品牌、固定製造商，以及系列、品類、顏色、尺寸的來源表對應欄位名稱。</p>
        
        <div class="input-row">
          <label style="flex:0 0 140px;">全域固定品牌 (選填):</label>
          <input type="text" id="cfg_brand_fixed" value="${escapeHtml(source.brand_fixed || '')}" placeholder="例如: RUKKA，未填寫則留空">
        </div>
        <div class="input-row">
          <label style="flex:0 0 140px;">全域固定製造商 (選填):</label>
          <input type="text" id="cfg_manufacturer_fixed" value="${escapeHtml(source.manufacturer_fixed || '')}" placeholder="未填寫且無設定時預設自動採用品牌">
        </div>
        <div class="input-row">
          <label style="flex:0 0 140px;">系列來源欄位 (Collection):</label>
          <select id="cfg_collection_column" class="form-select">${buildColSelectOptions(currentSheetHeaders, source.collection_column || 'COLLECTION')}</select>
        </div>
        <div class="input-row">
          <label style="flex:0 0 140px;">品類來源欄位 (Type):</label>
          <select id="cfg_type_column" class="form-select">${buildColSelectOptions(currentSheetHeaders, source.type_column || 'TYPE')}</select>
        </div>
        <div class="input-row">
          <label style="flex:0 0 140px;">顏色來源欄位 (Color):</label>
          <select id="cfg_color_column" class="form-select">${buildColSelectOptions(currentSheetHeaders, source.color_column || '中文顏色')}</select>
        </div>
        <div class="input-row">
          <label style="flex:0 0 140px;">尺寸來源欄位 (Size):</label>
          <select id="cfg_size_column" class="form-select">${buildColSelectOptions(currentSheetHeaders, source.size_column || 'SIZE')}</select>
        </div>
      </div>

      <div class="form-group">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
          <h4 style="margin:0;">全域固定欄位設定 (Global Fixed Mappings)</h4>
          <button type="button" class="btn btn-secondary btn-sm" onclick="window.addSourceFixedRow()"><span class="material-symbols-outlined" style="font-size:16px;">add</span> 新增全域固定欄位</button>
        </div>
        <p style="font-size:0.82rem; color:#64748b; margin-bottom:12px;">設定所有報價單通用的標準固定填寫內容（如包裝審核宣告、應稅、進口、法定種類、負責廠商等）。若個別品類模板有另外設定，將以模板設定優先。</p>
        <div id="sourceFixedContainer">${sourceFixedHtml}</div>
      </div>
    `;

    // Dynamic re-population of columns when cfg_sheet_name or cfg_header_row changes
    const cfgSheetSelect = document.getElementById('cfg_sheet_name');
    const cfgHeaderInput = document.getElementById('cfg_header_row');
    const updateColumnDropdowns = () => {
      if (!loadedWorkbook) return;
      const wsName = cfgSheetSelect ? cfgSheetSelect.value : source.sheet_name;
      const hRow = cfgHeaderInput ? (parseInt(cfgHeaderInput.value) || 3) : 3;
      const ws = loadedWorkbook.sheet(wsName) || loadedWorkbook.sheet(0);
      if (ws) {
        const headers = getSheetHeadersList(ws, hRow);
        if (headers.length > 0) {
          const updateSelect = (id, curVal, emptyLabel) => {
            const el = document.getElementById(id);
            if (el) el.innerHTML = buildColSelectOptions(headers, el.value || curVal, emptyLabel);
          };
          updateSelect('cfg_filter_column', source.filter_column || '', '-- 不啟用篩選 (處理全部列) --');
          updateSelect('cfg_collection_column', source.collection_column || 'COLLECTION');
          updateSelect('cfg_type_column', source.type_column || 'TYPE');
          updateSelect('cfg_color_column', source.color_column || '中文顏色');
          updateSelect('cfg_size_column', source.size_column || 'SIZE');
        }
      }
    };
    if (cfgSheetSelect) cfgSheetSelect.addEventListener('change', updateColumnDropdowns);
    if (cfgHeaderInput) cfgHeaderInput.addEventListener('input', updateColumnDropdowns);
  }

  function parseSourceConfig() {
    const fixed = {};
    document.querySelectorAll('#sourceFixedContainer .fixed-field-row').forEach(row => {
      const k = row.querySelector('.key-input')?.value.trim();
      const v = row.querySelector('.val-input')?.value.trim();
      if (k) fixed[k] = v;
    });

    const getVal = (id) => {
      const el = document.getElementById(id);
      return el ? el.value.trim() : '';
    };

    return {
      file_path: "商品資料.xlsx",
      sheet_name: getVal('cfg_sheet_name') || 'Sheet1',
      header_row: Math.max(1, parseInt(document.getElementById('cfg_header_row')?.value) || 3),
      row_start: Math.max(2, parseInt(document.getElementById('cfg_row_start')?.value) || 4),
      filter_column: getVal('cfg_filter_column'),
      brand_fixed: document.getElementById('cfg_brand_fixed')?.value.trim() || '',
      manufacturer_fixed: document.getElementById('cfg_manufacturer_fixed')?.value.trim() || '',
      collection_column: getVal('cfg_collection_column'),
      type_column: getVal('cfg_type_column'),
      color_column: getVal('cfg_color_column'),
      size_column: getVal('cfg_size_column'),
      fixed: fixed
    };
  }

  // Wizard Launch & Confirmation
  let pendingWizardKeywords = [];

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
          wizardScannedCount.textContent = scanResult.requiredColumns.length;

          // Initialize wizard keyword tag chips
          pendingWizardKeywords = [defaultName];
          const wizardKeywordsTagsContainer = document.getElementById('wizardKeywordsTagsContainer');
          const wizardKeywordInput = document.getElementById('wizardKeywordInput');
          const btnAddWizardKeyword = document.getElementById('btnAddWizardKeyword');

          function renderWizardKeywordsTags() {
            if (!wizardKeywordsTagsContainer) return;
            wizardKeywordsTagsContainer.innerHTML = '';
            if (pendingWizardKeywords.length === 0) {
              wizardKeywordsTagsContainer.innerHTML = '<span class="keyword-tag-empty">尚未設定關鍵字（輸入關鍵字後點擊 + 或按 Enter 新增）</span>';
              return;
            }
            pendingWizardKeywords.forEach((kw, index) => {
              const pill = document.createElement('div');
              pill.className = 'keyword-tag-pill';
              pill.innerHTML = `
                <span class="material-symbols-outlined">label</span>
                <span>${escapeHtml(kw)}</span>
                <button type="button" class="keyword-tag-remove" title="移除關鍵字" data-index="${index}">
                  <span class="material-symbols-outlined">close</span>
                </button>
              `;
              pill.querySelector('.keyword-tag-remove').addEventListener('click', (e) => {
                e.stopPropagation();
                pendingWizardKeywords.splice(index, 1);
                renderWizardKeywordsTags();
              });
              wizardKeywordsTagsContainer.appendChild(pill);
            });
          }

          function addWizardKeywordFromInput() {
            if (!wizardKeywordInput) return;
            const val = wizardKeywordInput.value.trim();
            if (!val) return;
            const items = val.split(/[,，\n]/).map(s => s.trim()).filter(Boolean);
            items.forEach(item => {
              if (!pendingWizardKeywords.includes(item)) {
                pendingWizardKeywords.push(item);
              }
            });
            wizardKeywordInput.value = '';
            renderWizardKeywordsTags();
          }

          if (btnAddWizardKeyword) {
            btnAddWizardKeyword.onclick = addWizardKeywordFromInput;
          }
          if (wizardKeywordInput) {
            wizardKeywordInput.onkeydown = (e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                addWizardKeywordFromInput();
              }
            };
          }
          renderWizardKeywordsTags();

          wizardTableBody.innerHTML = '';

          const availableHeaders = getAvailableSourceHeaders();
          const newCols = scanResult.requiredColumns.filter(c => c.status === 'new');
          const inheritedCols = scanResult.requiredColumns.filter(c => c.status === 'inherited');
          const systemCols = scanResult.requiredColumns.filter(c => c.status === 'system');

          const renderWizardGroup = (groupId, groupTitle, iconName, cols, defaultExpanded, groupTypeClass, emptyMsg) => {
            const headerTr = document.createElement('tr');
            headerTr.className = `wizard-group-header wizard-group-${groupTypeClass}`;

            let countBadge = '';
            if (groupTypeClass === 'new') {
              countBadge = `<span class="unmatched-count-badge" style="background:#fef3c7; color:#b45309; border-color:#fde68a;">${cols.length} 個待配置</span>`;
            } else if (groupTypeClass === 'inherited') {
              countBadge = `<span class="unmatched-count-badge" style="background:#dcfce7; color:#15803d; border-color:#bbf7d0;">${cols.length} 個已繼承</span>`;
            } else {
              countBadge = `<span class="unmatched-count-badge" style="background:#e0e7ff; color:#4338ca; border-color:#c7d2fe;">${cols.length} 個自動處理</span>`;
            }

            const toggleText = defaultExpanded ? `收合 (${cols.length})` : `展開 (${cols.length})`;
            const toggleExpandedClass = defaultExpanded ? 'expanded' : '';

            headerTr.innerHTML = `
              <td colspan="4">
                <div class="wizard-group-header-content">
                  <div class="wizard-group-title">
                    <span class="${iconName === 'auto_awesome' ? 'material-icons' : 'material-symbols-outlined'}" style="font-size: 1.2rem;">${iconName}</span>
                    <span>${groupTitle}</span>
                    ${countBadge}
                  </div>
                  <button type="button" class="wizard-toggle-btn ${toggleExpandedClass}" title="展開/收合">
                    <span class="material-symbols-outlined">expand_more</span>
                    <span class="wizard-toggle-label">${toggleText}</span>
                  </button>
                </div>
              </td>
            `;

            wizardTableBody.appendChild(headerTr);

            const rowElements = [];

            if (cols.length === 0 && emptyMsg) {
              const emptyTr = document.createElement('tr');
              emptyTr.className = `wizard-empty-group-row ${defaultExpanded ? '' : 'hidden'}`;
              emptyTr.innerHTML = `<td colspan="4">${emptyMsg}</td>`;
              wizardTableBody.appendChild(emptyTr);
              rowElements.push(emptyTr);
            } else {
              cols.forEach((col) => {
                const tr = document.createElement('tr');
                tr.className = `wizard-item-row wizard-row-${groupTypeClass} ${defaultExpanded ? '' : 'hidden'}`;
                tr.dataset.colName = col.colName;
                tr.dataset.status = col.status;

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
                  <td><strong>${escapeHtml(col.colName)}</strong></td>
                  <td>${statusBadge}</td>
                  <td>
                    <select class="wizard-type-select" ${isSys ? 'disabled' : ''}>
                      <option value="dynamic" ${col.mappingType === 'dynamic' ? 'selected' : ''}>動態對應 (來源表)</option>
                      <option value="fixed" ${col.mappingType === 'fixed' ? 'selected' : ''}>固定值 (固定內容)</option>
                      <option value="system" ${col.mappingType === 'system' ? 'selected' : ''}>系統自動處理</option>
                    </select>
                  </td>
                  <td>
                    <select class="wizard-val-select form-select ${col.mappingType !== 'dynamic' ? 'hidden' : ''}" ${isSys ? 'disabled' : ''}>
                      ${buildColSelectOptions(availableHeaders, col.defaultValue || '', '-- 請選擇來源表欄位 --')}
                    </select>
                    <input type="text" class="wizard-val-input ${col.mappingType !== 'fixed' ? 'hidden' : ''}" value="${escapeHtml(col.defaultValue || '')}" placeholder="固定填寫內容" ${isSys ? 'disabled' : ''}>
                    <input type="text" class="wizard-val-sys ${col.mappingType !== 'system' ? 'hidden' : ''}" value="(系統自動處理)" disabled>
                  </td>
                `;

                const select = tr.querySelector('.wizard-type-select');
                const valSelect = tr.querySelector('.wizard-val-select');
                const valInput = tr.querySelector('.wizard-val-input');
                const valSys = tr.querySelector('.wizard-val-sys');

                select.addEventListener('change', () => {
                  if (select.value === 'dynamic') {
                    if (valSelect) { valSelect.classList.remove('hidden'); valSelect.disabled = false; }
                    if (valInput) valInput.classList.add('hidden');
                    if (valSys) valSys.classList.add('hidden');
                  } else if (select.value === 'fixed') {
                    if (valSelect) valSelect.classList.add('hidden');
                    if (valInput) { valInput.classList.remove('hidden'); valInput.disabled = false; }
                    if (valSys) valSys.classList.add('hidden');
                  } else {
                    if (valSelect) valSelect.classList.add('hidden');
                    if (valInput) valInput.classList.add('hidden');
                    if (valSys) { valSys.classList.remove('hidden'); valSys.disabled = true; }
                  }
                });

                wizardTableBody.appendChild(tr);
                rowElements.push(tr);
              });
            }

            const toggleBtn = headerTr.querySelector('.wizard-toggle-btn');
            const toggleLabel = headerTr.querySelector('.wizard-toggle-label');

            headerTr.addEventListener('click', () => {
              const isExpanded = toggleBtn.classList.contains('expanded');
              if (isExpanded) {
                toggleBtn.classList.remove('expanded');
                rowElements.forEach(r => r.classList.add('hidden'));
                if (toggleLabel) toggleLabel.textContent = `展開 (${cols.length})`;
              } else {
                toggleBtn.classList.add('expanded');
                rowElements.forEach(r => r.classList.remove('hidden'));
                if (toggleLabel) toggleLabel.textContent = `收合 (${cols.length})`;
              }
            });
          };

          // 1. 新必填欄位 (置頂，預設展開)
          renderWizardGroup('new', '新必填欄位（需手動指定來源或固定值）', 'warning', newCols, true, 'new', '沒有新必填欄位，所有欄位均已自動匹配或繼承。');

          // 2. 繼承自預設 (預設折疊收合)
          renderWizardGroup('inherited', '繼承自預設（已自動沿用胸背帶對應設定）', 'auto_awesome', inheritedCols, false, 'inherited', '（無繼承欄位）');

          // 3. 系統自動處理 (預設折疊收合)
          renderWizardGroup('system', '系統自動處理（由生成引擎即時計算，無需設定）', 'smart_toy', systemCols, false, 'system', '（無系統自動處理欄位）');

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

      const dynamic = {};
      const fixed = {};

      const itemRows = wizardTableBody.querySelectorAll('tr.wizard-item-row');
      itemRows.forEach(tr => {
        const colName = tr.dataset.colName;
        const typeSelect = tr.querySelector('.wizard-type-select');
        const valSelect = tr.querySelector('.wizard-val-select');
        const valInput = tr.querySelector('.wizard-val-input');
        if (!colName || !typeSelect) return;

        const mType = typeSelect.value;
        if (mType === 'dynamic' && valSelect) {
          const val = valSelect.value.trim();
          if (val) dynamic[colName] = val;
        } else if (mType === 'fixed' && valInput) {
          const val = valInput.value.trim();
          if (val) fixed[colName] = val;
        }
      });

      const profileId = 'profile_' + Date.now();
      const keywords = pendingWizardKeywords.length > 0 ? [...pendingWizardKeywords] : [profName];

      const newProfile = {
        id: profileId,
        name: profName,
        keywords: keywords,
        template_type: profileId,
        template_id: profileId,
        template_file_name: pendingWizard.file.name,
        category_name: categoryName,
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

  // Open Config Modal and switch to target tab
  function openConfigModal(targetTabId = 'tabTemplateProfiles') {
    renderTemplateProfilesUI();
    renderSourceConfig(currentSourceConfig);
    configModal.classList.remove('hidden');

    tabBtns.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));

    const targetBtn = Array.from(tabBtns).find(b => b.dataset.target === targetTabId);
    if (targetBtn) {
      targetBtn.classList.add('active');
      document.getElementById(targetTabId)?.classList.add('active');
    }
  }

  // Modal for Unmatched Template Items Alert
  function showUnmatchedTemplateModal(unmatchedItems) {
    return new Promise((resolve) => {
      const modal = document.getElementById('unmatchedTemplateModal');
      const typeCountEl = document.getElementById('unmatchedTypeCount');
      const countEl = document.getElementById('unmatchedTotalCount');
      const listCountEl = document.getElementById('unmatchedListCount');
      const itemsListCountEl = document.getElementById('unmatchedItemsListCount');
      const tbody = document.getElementById('unmatchedTableBody');
      const btnGo = document.getElementById('btnGoToConfigTemplates');
      const btnIgnore = document.getElementById('btnIgnoreUnmatched');
      const btnClose = document.getElementById('unmatchedModalClose');

      if (!modal) {
        resolve({ action: 'ignore' });
        return;
      }

      // Group unmatched items by Type
      const groupMap = new Map();
      unmatchedItems.forEach(item => {
        const rawType = (item.type || '').toString().trim();
        const groupKey = rawType !== '' ? rawType : '(未填寫 / 空白)';
        if (!groupMap.has(groupKey)) {
          groupMap.set(groupKey, {
            type: groupKey,
            rawType: rawType,
            isUntyped: rawType === '',
            items: []
          });
        }
        groupMap.get(groupKey).items.push(item);
      });

      // Sort groups by count descending (most frequent first)
      const groups = Array.from(groupMap.values()).sort((a, b) => b.items.length - a.items.length);
      const totalItemsCount = unmatchedItems.length;
      const uniqueTypeCount = groups.length;

      if (typeCountEl) typeCountEl.textContent = uniqueTypeCount;
      if (countEl) countEl.textContent = totalItemsCount;
      if (listCountEl) listCountEl.textContent = uniqueTypeCount;
      if (itemsListCountEl) itemsListCountEl.textContent = totalItemsCount;

      if (tbody) {
        tbody.innerHTML = '';
        groups.forEach((group, idx) => {
          const mainTr = document.createElement('tr');
          mainTr.className = 'unmatched-type-row';

          const typeDisplay = group.isUntyped
            ? `<span class="unmatched-type-untyped"><em>(未填寫 / 空白)</em></span>`
            : `<strong>${escapeHtml(group.type)}</strong>`;

          mainTr.innerHTML = `
            <td>
              <div style="display: flex; align-items: center; gap: 8px;">
                <span class="material-symbols-outlined" style="font-size: 1.1rem; color: #f59e0b;">category</span>
                <span>${typeDisplay}</span>
              </div>
            </td>
            <td style="text-align: center;">
              <span class="unmatched-count-badge">${group.items.length} 筆</span>
            </td>
            <td style="text-align: right;">
              <button type="button" class="unmatched-toggle-btn" title="展開/收合商品明細">
                <span class="material-symbols-outlined">expand_more</span>
                <span class="toggle-text">展開明細 (${group.items.length})</span>
              </button>
            </td>
          `;

          // Detail row
          const detailTr = document.createElement('tr');
          detailTr.className = 'unmatched-detail-row hidden';

          const itemsHtml = group.items.map(item => `
            <div class="unmatched-sub-item">
              <span class="unmatched-row-badge">第 ${item.row} 列</span>
              <span class="unmatched-sub-item-name">${escapeHtml(item.name)}</span>
            </div>
          `).join('');

          detailTr.innerHTML = `
            <td colspan="3">
              <div class="unmatched-detail-container">
                <div style="margin-bottom: 6px; font-size: 0.8rem; font-weight: 600; color: #64748b;">包含以下商品（共 ${group.items.length} 筆）：</div>
                ${itemsHtml}
              </div>
            </td>
          `;

          const toggleAccordion = () => {
            const isHidden = detailTr.classList.contains('hidden');
            const toggleBtn = mainTr.querySelector('.unmatched-toggle-btn');
            const toggleText = mainTr.querySelector('.toggle-text');
            if (isHidden) {
              detailTr.classList.remove('hidden');
              mainTr.classList.add('is-expanded');
              if (toggleBtn) toggleBtn.classList.add('expanded');
              if (toggleText) toggleText.textContent = `收合明細 (${group.items.length})`;
            } else {
              detailTr.classList.add('hidden');
              mainTr.classList.remove('is-expanded');
              if (toggleBtn) toggleBtn.classList.remove('expanded');
              if (toggleText) toggleText.textContent = `展開明細 (${group.items.length})`;
            }
          };

          mainTr.addEventListener('click', () => {
            toggleAccordion();
          });

          tbody.appendChild(mainTr);
          tbody.appendChild(detailTr);
        });
      }

      modal.classList.remove('hidden');

      function cleanup() {
        modal.classList.add('hidden');
        if (btnGo) btnGo.removeEventListener('click', onGo);
        if (btnIgnore) btnIgnore.removeEventListener('click', onIgnore);
        if (btnClose) btnClose.removeEventListener('click', onIgnore);
      }

      function onGo() {
        cleanup();
        resolve({ action: 'go_to_config' });
      }

      function onIgnore() {
        cleanup();
        resolve({ action: 'ignore' });
      }

      if (btnGo) btnGo.addEventListener('click', onGo);
      if (btnIgnore) btnIgnore.addEventListener('click', onIgnore);
      if (btnClose) btnClose.addEventListener('click', onIgnore);
    });
  }

  // 來源 Excel 載入與智慧檢查兩步驟精靈 (2-Step Excel Upload Wizard)
  function showExcelUploadWizard(file, wb) {
    return new Promise((resolve) => {
      if (!excelUploadWizardModal) {
        resolve({ action: 'cancel' });
        return;
      }

      // Step 1: Initialize Sheets and Row settings
      if (uploadWizardFileName) {
        uploadWizardFileName.textContent = file.name;
      }

      const allSheets = wb.sheets ? wb.sheets() : [];
      let defaultWs = null;
      if (window.SharedUtils) {
        defaultWs = window.SharedUtils.getSourceSheet(wb, currentSourceConfig?.sheet_name, currentSourceConfig?.header_row || 3);
      } else {
        defaultWs = wb.sheet(currentSourceConfig?.sheet_name) || wb.sheet(0);
      }
      const defaultSheetName = defaultWs ? defaultWs.name() : (allSheets[0] ? allSheets[0].name() : 'Sheet1');

      if (wizardSheetSelect) {
        wizardSheetSelect.innerHTML = '';
        allSheets.forEach(s => {
          const opt = document.createElement('option');
          opt.value = s.name();
          const rCount = s.usedRange() ? s.usedRange().endCell().rowNumber() : 0;
          opt.textContent = `${s.name()} (${rCount > 0 ? `共 ${rCount} 列` : '空白工作表'})`;
          if (s.name() === defaultSheetName) {
            opt.selected = true;
          }
          wizardSheetSelect.appendChild(opt);
        });
      }

      if (wizardHeaderRow) wizardHeaderRow.value = currentSourceConfig?.header_row || 3;
      if (wizardRowStart) wizardRowStart.value = currentSourceConfig?.row_start || 4;

      function updateWizardStep1FilterOptions() {
        const selSheetName = wizardSheetSelect ? wizardSheetSelect.value : defaultSheetName;
        const hRow = parseInt(wizardHeaderRow?.value) || 3;
        const targetWs = wb.sheet(selSheetName);
        let headers = [];
        if (targetWs) {
          headers = getSheetHeadersList(targetWs, hRow);
        }
        if (headers.length === 0) {
          headers = getCommonSourceHeaders();
        }
        if (wizardFilterColumn) {
          const curVal = wizardFilterColumn.value || currentSourceConfig?.filter_column || '';
          wizardFilterColumn.innerHTML = buildColSelectOptions(headers, curVal, '-- 不啟用篩選 (處理全部列) --');
        }
      }

      updateWizardStep1FilterOptions();
      if (wizardSheetSelect) wizardSheetSelect.addEventListener('change', updateWizardStep1FilterOptions);
      if (wizardHeaderRow) wizardHeaderRow.addEventListener('input', updateWizardStep1FilterOptions);

      // Set Stepper to Step 1
      function setStepperView(step) {
        if (step === 1) {
          if (wizardStepIndicator1) wizardStepIndicator1.className = 'stepper-step active';
          if (wizardStepIndicator2) wizardStepIndicator2.className = 'stepper-step';
          if (wizardStepperLine) wizardStepperLine.className = 'stepper-line';
          if (uploadWizardStep1) uploadWizardStep1.classList.remove('hidden');
          if (uploadWizardStep2) uploadWizardStep2.classList.add('hidden');
          if (btnWizardPrev) btnWizardPrev.classList.add('hidden');
          if (btnWizardNext) btnWizardNext.classList.remove('hidden');
          if (btnWizardConfirm) btnWizardConfirm.classList.add('hidden');
        } else if (step === 2) {
          if (wizardStepIndicator1) wizardStepIndicator1.className = 'stepper-step completed';
          if (wizardStepIndicator2) wizardStepIndicator2.className = 'stepper-step active';
          if (wizardStepperLine) wizardStepperLine.className = 'stepper-line active';
          if (uploadWizardStep1) uploadWizardStep1.classList.add('hidden');
          if (uploadWizardStep2) uploadWizardStep2.classList.remove('hidden');
          if (btnWizardPrev) btnWizardPrev.classList.remove('hidden');
          if (btnWizardNext) btnWizardNext.classList.add('hidden');
          if (btnWizardConfirm) btnWizardConfirm.classList.remove('hidden');
        }
      }

      setStepperView(1);
      excelUploadWizardModal.classList.remove('hidden');

      let currentStep2Data = {
        ws: null,
        totalR: 0,
        selectedSheetName: defaultSheetName,
        selectedHeaderRow: 3,
        selectedRowStart: 4,
        selectedFilterCol: currentSourceConfig?.filter_column || '',
        detectedProfiles: [],
        missingItems: [],
        excelHeaderNames: [],
        unmatchedItems: [],
        unmatchedGroups: []
      };

      function cleanup() {
        excelUploadWizardModal.classList.add('hidden');
        if (btnWizardNext) btnWizardNext.removeEventListener('click', onNext);
        if (btnWizardPrev) btnWizardPrev.removeEventListener('click', onPrev);
        if (btnWizardConfirm) btnWizardConfirm.removeEventListener('click', onConfirm);
        if (btnWizardCancel) btnWizardCancel.removeEventListener('click', onCancel);
        if (uploadWizardClose) uploadWizardClose.removeEventListener('click', onCancel);
        if (btnWizardGoToConfig) btnWizardGoToConfig.removeEventListener('click', onGoToConfig);
        if (wizardSheetSelect) wizardSheetSelect.removeEventListener('change', updateWizardStep1FilterOptions);
        if (wizardHeaderRow) wizardHeaderRow.removeEventListener('input', updateWizardStep1FilterOptions);
      }

      // Step 2 Live Template Detection & Unmatched Type Re-calculation
      function updateStep2TemplateDetection() {
        if (!currentStep2Data.ws) return;

        const ws = currentStep2Data.ws;
        const selectedSheetName = currentStep2Data.selectedSheetName;
        const selectedHeaderRow = currentStep2Data.selectedHeaderRow;
        const selectedRowStart = currentStep2Data.selectedRowStart;
        const totalR = currentStep2Data.totalR;

        const selectedFilterCol = wizardFilterColumn ? wizardFilterColumn.value.trim() : (currentSourceConfig?.filter_column || '');
        const selectedTypeCol = wizardTypeColumn ? wizardTypeColumn.value.trim() : (currentSourceConfig?.type_column || 'TYPE');
        const selectedCollecCol = wizardCollectionColumn ? wizardCollectionColumn.value.trim() : (currentSourceConfig?.collection_column || 'COLLECTION');
        const selectedColorCol = wizardColorColumn ? wizardColorColumn.value.trim() : (currentSourceConfig?.color_column || '中文顏色');
        const selectedSizeCol = wizardSizeColumn ? wizardSizeColumn.value.trim() : (currentSourceConfig?.size_column || 'SIZE');

        const tempSourceConfig = {
          ...currentSourceConfig,
          sheet_name: selectedSheetName,
          header_row: selectedHeaderRow,
          row_start: selectedRowStart,
          filter_column: selectedFilterCol,
          type_column: selectedTypeCol,
          collection_column: selectedCollecCol,
          color_column: selectedColorCol,
          size_column: selectedSizeCol
        };

        // Clone template profiles and overlay any mappings currently chosen in mappingTableBody
        const tempProfiles = templateProfiles.map(p => ({
          ...p,
          field_mappings: {
            dynamic: { ...(p.field_mappings?.dynamic || {}) },
            fixed: { ...(p.field_mappings?.fixed || {}) }
          }
        }));

        if (mappingTableBody && currentStep2Data.missingItems) {
          const rows = mappingTableBody.querySelectorAll('tr');
          rows.forEach((r, idx) => {
            const item = currentStep2Data.missingItems[idx];
            if (!item) return;
            const select = r.querySelector('.mapping-col-select');
            const val = select ? select.value : '';
            if (val && val !== '__SKIP__' && val !== '__FIXED__') {
              for (const prof of item.profiles || []) {
                const tp = tempProfiles.find(p => p.id === prof.id);
                if (tp) {
                  tp.field_mappings.dynamic[item.targetField] = val;
                }
              }
            }
          });
        }

        const detectResult = detectRequiredTemplates(ws, selectedHeaderRow, selectedRowStart, totalR, tempProfiles, tempSourceConfig);
        const detectedProfiles = Array.isArray(detectResult) ? detectResult : (detectResult.detectedProfiles || tempProfiles);
        const unmatchedItems = Array.isArray(detectResult) ? [] : (detectResult.unmatchedItems || []);

        // Aggregate unmatched items by Type
        const typeCountMap = new Map();
        unmatchedItems.forEach(item => {
          const rawType = (item.type || '').toString().trim();
          const t = rawType !== '' ? rawType : '(未填寫 / 空白)';
          if (!typeCountMap.has(t)) {
            typeCountMap.set(t, { type: t, rawType, isUntyped: rawType === '', items: [] });
          }
          typeCountMap.get(t).items.push(item);
        });
        const unmatchedGroups = Array.from(typeCountMap.values()).sort((a, b) => b.items.length - a.items.length);
        lastDetectedUnmatchedTypes = unmatchedGroups.map(g => ({ type: g.type, count: g.items.length }));

        currentStep2Data.detectedProfiles = detectedProfiles;
        currentStep2Data.unmatchedItems = unmatchedItems;
        currentStep2Data.unmatchedGroups = unmatchedGroups;

        // 1. Re-render detected templates
        if (detectedTemplatesContainer) {
          detectedTemplatesContainer.innerHTML = '';
          detectedProfiles.forEach(p => {
            const badge = document.createElement('span');
            badge.className = 'badge-template';
            badge.innerHTML = `<span class="material-symbols-outlined" style="font-size: 0.85rem;">category</span><span>${escapeHtml(p.name)}</span>`;
            detectedTemplatesContainer.appendChild(badge);
          });
        }

        // 2. Re-render unmatched alert box
        if (wizardUnmatchedAlertBox) {
          if (unmatchedItems.length > 0) {
            wizardUnmatchedAlertBox.classList.remove('hidden');
            if (wizardUnmatchedTypeCount) wizardUnmatchedTypeCount.textContent = unmatchedGroups.length;
            if (wizardUnmatchedTotalCount) wizardUnmatchedTotalCount.textContent = unmatchedItems.length;

            if (wizardUnmatchedAccordionList) {
              wizardUnmatchedAccordionList.innerHTML = '';
              unmatchedGroups.forEach(g => {
                const itemDiv = document.createElement('div');
                itemDiv.className = 'wizard-unmatched-accordion-item';

                const headDiv = document.createElement('div');
                headDiv.className = 'wizard-unmatched-type-head';
                headDiv.innerHTML = `
                  <div style="display:flex; align-items:center; gap:6px;">
                    <span class="material-symbols-outlined" style="font-size:0.95rem; color:#f59e0b;">label</span>
                    <span>${g.isUntyped ? '<em>(未填寫 / 空白)</em>' : `<strong>${escapeHtml(g.type)}</strong>`}</span>
                  </div>
                  <div style="display:flex; align-items:center; gap:6px;">
                    <span class="unmatched-count-badge">${g.items.length} 筆</span>
                    <span class="material-symbols-outlined toggle-icon" style="font-size:1.1rem; transition: transform 0.2s;">expand_more</span>
                  </div>
                `;

                const detailDiv = document.createElement('div');
                detailDiv.className = 'wizard-unmatched-detail-pane hidden';
                const itemsHtml = g.items.map(it => `
                  <div class="unmatched-sub-item">
                    <span class="unmatched-row-badge">第 ${it.row} 列</span>
                    <span class="unmatched-sub-item-name">${escapeHtml(it.name)}</span>
                  </div>
                `).join('');
                detailDiv.innerHTML = itemsHtml;

                headDiv.addEventListener('click', () => {
                  const isHidden = detailDiv.classList.contains('hidden');
                  const toggleIcon = headDiv.querySelector('.toggle-icon');
                  if (isHidden) {
                    detailDiv.classList.remove('hidden');
                    if (toggleIcon) toggleIcon.style.transform = 'rotate(180deg)';
                  } else {
                    detailDiv.classList.add('hidden');
                    if (toggleIcon) toggleIcon.style.transform = 'rotate(0deg)';
                  }
                });

                itemDiv.appendChild(headDiv);
                itemDiv.appendChild(detailDiv);
                wizardUnmatchedAccordionList.appendChild(itemDiv);
              });
            }
          } else {
            wizardUnmatchedAlertBox.classList.add('hidden');
          }
        }
      }

      function onNext() {
        const selectedSheetName = wizardSheetSelect ? wizardSheetSelect.value : defaultSheetName;
        const selectedHeaderRow = Math.max(1, parseInt(wizardHeaderRow?.value) || 3);
        const selectedRowStart = Math.max(selectedHeaderRow + 1, parseInt(wizardRowStart?.value) || (selectedHeaderRow + 1));
        const selectedFilterCol = wizardFilterColumn ? wizardFilterColumn.value.trim() : '';

        const ws = wb.sheet(selectedSheetName);
        if (!ws || !ws.usedRange()) {
          alert(`工作表「${selectedSheetName}」中找不到任何有效資料，請重新選擇工作表！`);
          return;
        }

        const totalR = ws.usedRange().endCell().rowNumber();
        if (selectedHeaderRow >= totalR) {
          alert(`標題列設定為第 ${selectedHeaderRow} 列，已超出或等於工作表總列數 (${totalR} 列)，請確認標題列位置！`);
          return;
        }
        if (selectedRowStart > totalR) {
          alert(`資料起始列設定為第 ${selectedRowStart} 列，已超出工作表總列數 (${totalR} 列)，請確認起始列位置！`);
          return;
        }

        const tempSourceConfig = {
          ...currentSourceConfig,
          sheet_name: selectedSheetName,
          header_row: selectedHeaderRow,
          row_start: selectedRowStart,
          filter_column: selectedFilterCol
        };

        // Check missing columns
        const checkResult = checkMissingColumns(templateProfiles, ws, selectedHeaderRow, tempSourceConfig);
        const missingItems = checkResult.missingItems || [];
        const excelHeaderNames = checkResult.excelHeaderNames || [];

        currentStep2Data = {
          ws,
          totalR,
          selectedSheetName,
          selectedHeaderRow,
          selectedRowStart,
          selectedFilterCol,
          detectedProfiles: [],
          unmatchedItems: [],
          unmatchedGroups: [],
          missingItems,
          excelHeaderNames
        };

        // 3. Global Attributes & Dynamic Mappings
        if (wizardBrandFixed) wizardBrandFixed.value = currentSourceConfig?.brand_fixed || '';
        if (wizardManufacturerFixed) wizardManufacturerFixed.value = currentSourceConfig?.manufacturer_fixed || '';

        function populateColSelect(selEl, currentVal, fallbackKeys = []) {
          if (!selEl) return;
          selEl.innerHTML = '';
          const emptyOpt = document.createElement('option');
          emptyOpt.value = '';
          emptyOpt.textContent = '-- 請選擇來源欄位 (未指定則留空) --';
          selEl.appendChild(emptyOpt);

          let hasSelected = false;
          const directMatchHeader = currentVal ? excelHeaderNames.find(h => h.toUpperCase() === currentVal.toUpperCase()) : null;
          const fallbackMatchHeader = !directMatchHeader ? excelHeaderNames.find(h => fallbackKeys.some(k => k.toUpperCase() === h.toUpperCase())) : null;

          excelHeaderNames.forEach(h => {
            const opt = document.createElement('option');
            opt.value = h;
            const isDirect = directMatchHeader && (h.toUpperCase() === directMatchHeader.toUpperCase());
            const isFallback = !directMatchHeader && fallbackMatchHeader && (h.toUpperCase() === fallbackMatchHeader.toUpperCase());
            if (isDirect || isFallback) {
              opt.textContent = `${h} (推薦比對)`;
              if (!hasSelected) {
                opt.selected = true;
                hasSelected = true;
              }
            } else {
              opt.textContent = h;
            }
            selEl.appendChild(opt);
          });
        }

        populateColSelect(wizardCollectionColumn, currentSourceConfig?.collection_column, ['COLLECTION', '系列', 'Collection', '系列名稱']);
        populateColSelect(wizardTypeColumn, currentSourceConfig?.type_column, ['TYPE', '品類', '種類', '款式', 'Type', '類別']);
        populateColSelect(wizardColorColumn, currentSourceConfig?.color_column, ['中文顏色', 'COLOR', '顏色', 'Color', '顏色名稱']);
        populateColSelect(wizardSizeColumn, currentSourceConfig?.size_column, ['SIZE', '尺寸', 'Size', '規格尺寸']);

        // 4. Missing columns list
        if (missingColumnCount) missingColumnCount.textContent = missingItems.length;

        if (missingItems.length > 0) {
          if (missingColumnsContainer) missingColumnsContainer.classList.remove('hidden');
          if (allColumnsMatchedNotice) allColumnsMatchedNotice.classList.add('hidden');

          if (mappingTableBody) {
            mappingTableBody.innerHTML = '';
            missingItems.forEach(item => {
              const tr = document.createElement('tr');

              const tdTarget = document.createElement('td');
              tdTarget.innerHTML = `<span class="mapping-target-name">${escapeHtml(item.targetField)}</span>`;

              const tdProfiles = document.createElement('td');
              tdProfiles.textContent = item.profiles.map(p => p.name).join('、');

              const tdExpected = document.createElement('td');
              tdExpected.innerHTML = `<span class="mapping-expected-name">${escapeHtml(item.expectedSourceCol)}</span>`;

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

              const isMandatoryName = (item.key === '商品名稱' || item.key === '中文品名');

              if (!item.isFilter && !isMandatoryName) {
                const skipOpt = document.createElement('option');
                skipOpt.value = '__SKIP__';
                skipOpt.textContent = '【留空 / 略過此欄位】';
                select.appendChild(skipOpt);

                const fixedOpt = document.createElement('option');
                fixedOpt.value = '__FIXED__';
                fixedOpt.textContent = '【手動輸入固定值...】';
                select.appendChild(fixedOpt);
              } else if (isMandatoryName) {
                const fixedOpt = document.createElement('option');
                fixedOpt.value = '__FIXED__';
                fixedOpt.textContent = '【手動指定固定品名...】';
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
                updateStep2TemplateDetection();
              });

              tdSelect.appendChild(box);
              tr.appendChild(tdTarget);
              tr.appendChild(tdProfiles);
              tr.appendChild(tdExpected);
              tr.appendChild(tdSelect);
              mappingTableBody.appendChild(tr);
            });
          }
        } else {
          if (missingColumnsContainer) missingColumnsContainer.classList.add('hidden');
          if (allColumnsMatchedNotice) allColumnsMatchedNotice.classList.remove('hidden');
        }

        // Attach change listeners to dynamic mappings for real-time re-detection
        if (wizardTypeColumn) wizardTypeColumn.onchange = updateStep2TemplateDetection;
        if (wizardCollectionColumn) wizardCollectionColumn.onchange = updateStep2TemplateDetection;
        if (wizardColorColumn) wizardColorColumn.onchange = updateStep2TemplateDetection;
        if (wizardSizeColumn) wizardSizeColumn.onchange = updateStep2TemplateDetection;

        // Perform initial detection for Step 2
        updateStep2TemplateDetection();

        setStepperView(2);
      }

      function onPrev() {
        setStepperView(1);
      }

      function onGoToConfig() {
        cleanup();
        resolve({ action: 'go_to_config' });
      }

      function onCancel() {
        cleanup();
        resolve({ action: 'cancel' });
      }

      function onConfirm() {
        const mappings = {};
        let hasUnselected = false;

        if (currentStep2Data.missingItems && currentStep2Data.missingItems.length > 0 && mappingTableBody) {
          const rows = mappingTableBody.querySelectorAll('tr');
          rows.forEach((r, idx) => {
            const item = currentStep2Data.missingItems[idx];
            const select = r.querySelector('.mapping-col-select');
            const fixedInput = r.querySelector('.mapping-fixed-input');
            const val = select ? select.value : '';

            if (!val) {
              hasUnselected = true;
              if (select) select.style.borderColor = '#ef4444';
            } else {
              if (select) select.style.borderColor = '';
            }

            if (val === '__SKIP__') {
              mappings[item.key] = { type: 'skip', value: '', isFilter: item.isFilter, profiles: item.profiles };
            } else if (val === '__FIXED__') {
              mappings[item.key] = { type: 'fixed', value: (fixedInput ? fixedInput.value : '').trim(), isFilter: item.isFilter, profiles: item.profiles };
            } else {
              mappings[item.key] = { type: 'column', value: val, isFilter: item.isFilter, profiles: item.profiles };
            }
          });
        }

        if (hasUnselected) {
          alert('請為所有缺漏欄位選取對應來源、或選擇「留空」/「手動輸入固定值」！');
          return;
        }

        const remember = chkRememberMappings ? chkRememberMappings.checked : true;

        const resultData = {
          action: 'confirm',
          sheet_name: currentStep2Data.selectedSheetName,
          header_row: currentStep2Data.selectedHeaderRow,
          row_start: currentStep2Data.selectedRowStart,
          filter_column: currentStep2Data.selectedFilterCol,
          brand_fixed: wizardBrandFixed ? wizardBrandFixed.value.trim() : '',
          manufacturer_fixed: wizardManufacturerFixed ? wizardManufacturerFixed.value.trim() : '',
          collection_column: wizardCollectionColumn ? wizardCollectionColumn.value.trim() : '',
          type_column: wizardTypeColumn ? wizardTypeColumn.value.trim() : '',
          color_column: wizardColorColumn ? wizardColorColumn.value.trim() : '',
          size_column: wizardSizeColumn ? wizardSizeColumn.value.trim() : '',
          mappings,
          remember,
          detectedProfiles: currentStep2Data.detectedProfiles,
          unmatchedItems: currentStep2Data.unmatchedItems,
          totalRows: currentStep2Data.totalR,
          ws: currentStep2Data.ws
        };

        cleanup();
        resolve(resultData);
      }

      if (btnWizardNext) btnWizardNext.addEventListener('click', onNext);
      if (btnWizardPrev) btnWizardPrev.addEventListener('click', onPrev);
      if (btnWizardConfirm) btnWizardConfirm.addEventListener('click', onConfirm);
      if (btnWizardCancel) btnWizardCancel.addEventListener('click', onCancel);
      if (uploadWizardClose) uploadWizardClose.addEventListener('click', onCancel);
      if (btnWizardGoToConfig) btnWizardGoToConfig.addEventListener('click', onGoToConfig);
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
        }
        if (typeof parseGuiCollection === 'function') {
          const ca = parseGuiCollection();
          localStorage.setItem('coupang_collection_aliases', JSON.stringify(ca));
        }
        if (typeof parseGuiColor === 'function') {
          const cla = parseGuiColor();
          localStorage.setItem('coupang_color_aliases', JSON.stringify(cla));
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
          renderTemplateProfilesUI();
          renderSourceConfig(currentSourceConfig);

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
    openConfigModal('tabTemplateProfiles');
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

      const globalConfig = window.AppConfig.get();
      globalConfig.source = currentSourceConfig;
      localStorage.setItem('coupang_config', JSON.stringify(globalConfig));

      updateRangeHintUI();
      alert('所有設定檔已儲存！');
      configModal.classList.add('hidden');
    } catch (e) {
      alert('儲存設定時發生錯誤: ' + e.message);
    }
  });

  btnResetConfig.addEventListener('click', async () => {
    if (confirm('確定要還原為系統預設值嗎？自訂模板設定檔將被清除。')) {
      localStorage.removeItem('coupang_config');
      localStorage.removeItem('coupang_category_rules');
      localStorage.removeItem('coupang_templates');
      localStorage.removeItem('coupang_template_profiles');

      templateProfiles = window.AppConfig.getDefaultProfiles();
      for (const p of templateProfiles) {
        await window.StorageUtils.saveProfile(p);
      }

      currentSourceConfig = window.AppConfig.getDefaultSourceConfig();

      renderTemplateProfilesUI();
      renderSourceConfig(currentSourceConfig);

      if (inputRowStart) inputRowStart.value = currentSourceConfig.row_start || 4;
      if (inputRowEnd) inputRowEnd.value = "";
      updateRangeHintUI();
      alert('已還原為系統預設值。');
    }
  });

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
      photoFilesArray = Array.from(e.target.files).filter(file => {
        const p = (file.webkitRelativePath || file.name || '').replace(/\\/g, '/');
        return !/(?:^|\/)(result|output|__macosx|\.git)(?:\/|$)/i.test(p);
      });
      photoDirZone.classList.add('has-file');
      photoDirInfo.textContent = `已選取資料夾，共包含 ${photoFilesArray.length} 個商品圖檔案`;
      photoDirInfo.classList.remove('hidden');
      logMessage(`已載入照片資料夾，共包含 ${photoFilesArray.length} 個有效圖檔`);
      checkReady();
    }
  });

  function getSourceSheet(wb) {
    if (window.SharedUtils) {
      return window.SharedUtils.getSourceSheet(wb, currentSourceConfig?.sheet_name, currentSourceConfig?.header_row || 3);
    }
    if (!wb) return null;
    return wb.sheet(currentSourceConfig?.sheet_name) || wb.sheet(0);
  }

  function getValidRowsInfo(ws, customStart = null, customEnd = null) {
    if (!ws) return { validRows: [], totalDataRows: 0, startRow: 0, endRow: 0, isRangeInvalid: false };
    const headerRow = (currentSourceConfig?.header_row !== undefined) ? currentSourceConfig.header_row : 3;
    const defaultStartRow = (currentSourceConfig?.row_start !== undefined) ? currentSourceConfig.row_start : 4;
    const filterColName = currentSourceConfig?.filter_column;
    
    const range = ws.usedRange();
    if (!range) return { validRows: [], totalDataRows: 0, startRow: 0, endRow: 0, isRangeInvalid: false };
    const maxRow = range.endCell().rowNumber();

    const parsedStart = parseInt(customStart);
    const parsedEnd = parseInt(customEnd);

    const actualStart = Math.max(headerRow + 1, !isNaN(parsedStart) ? parsedStart : defaultStartRow);
    const actualEnd = (!isNaN(parsedEnd) && parsedEnd > 0) ? Math.min(maxRow, parsedEnd) : maxRow;

    if (actualStart > actualEnd) {
      return { validRows: [], totalDataRows: 0, startRow: actualStart, endRow: actualEnd, maxRow, isRangeInvalid: true };
    }

    const headers = buildHeaderMap(ws, headerRow);
    const filterColIdx = (filterColName && filterColName.trim() !== '') ? findHeaderColIdx(headers, filterColName) : null;

    const validRows = [];
    let totalDataRows = 0;

    for (let r = actualStart; r <= actualEnd; r++) {
      totalDataRows++;

      if (filterColIdx) {
        const filterVal = getCellValue(ws.cell(r, filterColIdx));
        if (filterVal === null || filterVal === undefined || String(filterVal).trim() === '') {
          continue;
        }
      }

      validRows.push(r);
    }

    return { validRows, totalDataRows, startRow: actualStart, endRow: actualEnd, maxRow, isRangeInvalid: false };
  }

  function calculateValidProducts(ws) {
    const sStart = inputRowStart ? inputRowStart.value : null;
    const sEnd = inputRowEnd ? inputRowEnd.value : null;
    const info = getValidRowsInfo(ws, sStart, sEnd);
    return { count: info.validRows.length, rows: info.validRows, totalDataRows: info.totalDataRows, maxRow: info.maxRow, isRangeInvalid: info.isRangeInvalid };
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

    if (info.isRangeInvalid) {
      rangeInfoHint.textContent = `[錯誤] 起始列 (第 ${info.startRow} 列) 不得大於結束列 (第 ${info.endRow} 列)！`;
      rangeInfoHint.className = 'range-hint warning';
      return;
    }

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
      if (!wb || !wb.sheets || wb.sheets().length === 0) {
        throw new Error('來源 Excel 中找不到任何可用的工作表！');
      }

      const wizardRes = await showExcelUploadWizard(file, wb);

      if (wizardRes.action === 'cancel') {
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

      if (wizardRes.action === 'go_to_config') {
        sourceExcelFile = null;
        loadedWorkbook = null;
        excelDropZone.classList.remove('has-file');
        excelFileInfo.classList.add('hidden');
        excelFileInfo.textContent = '';
        checkReady();
        updateRangeHintUI();
        openConfigModal('tabTemplateProfiles');
        logMessage('已開啟「對照表設定 ➔ 模板與設定檔」，請新增對應品類模板與關鍵字後，再重新載入 Excel。', 'info');
        return;
      }

      // Apply Source Config from Wizard
      currentSourceConfig.sheet_name = wizardRes.sheet_name;
      currentSourceConfig.header_row = wizardRes.header_row;
      currentSourceConfig.row_start = wizardRes.row_start;
      currentSourceConfig.filter_column = wizardRes.filter_column;
      currentSourceConfig.brand_fixed = wizardRes.brand_fixed;
      currentSourceConfig.manufacturer_fixed = wizardRes.manufacturer_fixed;
      currentSourceConfig.collection_column = wizardRes.collection_column;
      currentSourceConfig.type_column = wizardRes.type_column;
      currentSourceConfig.color_column = wizardRes.color_column;
      currentSourceConfig.size_column = wizardRes.size_column;

      // Apply missing column mappings
      let updatedCount = 0;
      for (const [key, mapping] of Object.entries(wizardRes.mappings || {})) {
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

      if (wizardRes.remember) {
        for (const p of templateProfiles) {
          await window.StorageUtils.saveProfile(p);
        }
        const globalConfig = window.AppConfig.get();
        globalConfig.source = currentSourceConfig;
        localStorage.setItem('coupang_config', JSON.stringify(globalConfig));
        if (typeof renderTemplateProfilesUI === 'function') renderTemplateProfilesUI();
        if (typeof renderSourceConfig === 'function') renderSourceConfig(currentSourceConfig);
      }

      sourceExcelFile = file;
      loadedWorkbook = wb;
      excelDropZone.classList.add('has-file');
      excelFileInfo.textContent = `已載入: ${file.name}（工作表: ${wizardRes.sheet_name}）`;
      excelFileInfo.classList.remove('hidden');
      checkReady();

      if (inputRowStart) {
        inputRowStart.value = currentSourceConfig.row_start || 4;
      }
      if (inputRowEnd) {
        inputRowEnd.placeholder = `最大 ${wizardRes.totalRows} 列`;
      }
      updateRangeHintUI();

      const stats = calculateValidProducts(wizardRes.ws);
      const tmplNames = (wizardRes.detectedProfiles || []).map(p => p.name).join('、');
      logMessage(`已成功載入來源 Excel: ${file.name}（工作表「${wizardRes.sheet_name}」共 ${wizardRes.totalRows} 列，自動套用模板：【${tmplNames}】，符合條件之有效商品共 ${stats.count} 筆）`, 'success');

      if (wizardRes.unmatchedItems && wizardRes.unmatchedItems.length > 0) {
        logMessage(`[提示] 來源 Excel 中有 ${wizardRes.unmatchedItems.length} 筆商品未匹配任何品類模板，已在對照表標註缺少的 Type。`, 'warning');
      }

      if (updatedCount > 0) {
        logMessage(`[欄位智慧補全] 已成功補齊/更新 ${updatedCount} 個欄位對應關係！`, 'success');
      }
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

    if (statTotalSku) statTotalSku.textContent = '0';
    if (statMatchedImages) statMatchedImages.textContent = '0';
    if (statMissingImages) statMissingImages.textContent = '0';
    if (statMissingBackLabels) statMissingBackLabels.textContent = '0';
    if (statUnmatchedTemplates) statUnmatchedTemplates.textContent = '0';

    processedResults = [];
    generatedExcelFiles = new Map();
    filesToExport = {};

    try {
      const arrayBuffer = await sourceExcelFile.arrayBuffer();
      const wb = await XlsxPopulate.fromDataAsync(arrayBuffer);
      const ws = getSourceSheet(wb);
      if (!ws) throw new Error('來源 Excel 中找不到任何可用的工作表！');

      const headerRow = (currentSourceConfig?.header_row !== undefined) ? currentSourceConfig.header_row : 3;
      const rowStartInput = inputRowStart ? inputRowStart.value : null;
      const rowEndInput = inputRowEnd ? inputRowEnd.value : null;
      const rangeInfo = getValidRowsInfo(ws, rowStartInput, rowEndInput);

      if (rangeInfo.isRangeInvalid) {
        alert(`處理範圍設定有誤：起始列 (第 ${rangeInfo.startRow} 列) 不得大於結束列 (第 ${rangeInfo.endRow} 列)！請修正後再試。`);
        setProgress(0, '處理範圍設定有誤');
        return;
      }

      const validRowIndices = rangeInfo.validRows;
      const rowStart = rangeInfo.startRow;
      const rowEnd = rangeInfo.endRow;
      const totalValidCount = validRowIndices.length;
      const totalRows = rangeInfo.maxRow;

      const headers = buildHeaderMap(ws, headerRow);
      let nameColIdx = null;
      for (const p of templateProfiles) {
        const dynName = p.field_mappings?.dynamic?.['商品名稱'] || p.field_mappings?.dynamic?.['中文品名'] || p.field_mappings?.dynamic?.['產品中文名稱'];
        if (dynName) {
          nameColIdx = findHeaderColIdx(headers, dynName);
          if (nameColIdx) break;
        }
      }
      if (!nameColIdx) {
        const nameCandidates = ['商品名稱', '中文品名', '產品中文名稱', '產品名稱', '商品中文名稱', '中文名稱', '品名(中)', '品名', '商品名', '商品', 'NAME', 'Item Name', 'Product Name', 'Description'];
        for (const cand of nameCandidates) {
          nameColIdx = findHeaderColIdx(headers, cand);
          if (nameColIdx) break;
        }
      }

      const filterColName = (currentSourceConfig?.filter_column !== undefined) ? currentSourceConfig.filter_column : '';
      const filterColIdx = (filterColName && filterColName.trim() !== '') ? findHeaderColIdx(headers, filterColName) : null;

      if (!nameColIdx) {
        throw new Error(`來源 Excel 表頭（第 ${headerRow} 列）缺少「商品名稱」或「中文品名」必要欄位！請在對照表設定或上傳時指定品名欄位。`);
      }

      logMessage(`處理範圍：第 ${rowStart} 列 至 第 ${rowEnd} 列（共掃描 ${rangeInfo.totalDataRows} 列，其中有效商品共 ${totalValidCount} 筆，工作表總列數: ${totalRows} 列）`);

      if (totalValidCount === 0) {
        setProgress(100, '所選範圍內無符合條件之有效商品資料。');
        logMessage('未找到任何符合篩選條件的有效商品資料。', 'warning');
        return;
      }

      const processor = new window.CoupangProcessor(
        photoFilesArray,
        arrayBuffer,
        templateProfiles
      );

      // Map to hold loaded Template Workbooks: Map<templateKey, { wb, ws, headerMap, nextRowIdx, profile, targetSubfolder, fileName, count }>
      const activeWorkbooks = new Map();

      async function getOrInitWorkbook(profileOrId, categoryHint, subfolderHint) {
        if (!Array.isArray(templateProfiles) || templateProfiles.length === 0) {
          templateProfiles = window.AppConfig.getDefaultProfiles();
        }

        let profile = null;
        if (profileOrId && typeof profileOrId === 'object' && profileOrId.name) {
          profile = profileOrId;
        } else if (typeof profileOrId === 'string') {
          profile = templateProfiles.find(p => p.id === profileOrId || p.template_type === profileOrId);
        }
        if (!profile) {
          profile = templateProfiles.find(p => p.id === 'LEASH') || templateProfiles[0] || window.AppConfig.getDefaultProfiles()[0];
        }

        const targetSubfolder = subfolderHint || (window.SharedUtils ? window.SharedUtils.getTemplateSubfolder(profile, templateProfiles) : (profile.name || '未分類品項'));
        const templateFileName = profile.template_file_name || (profile.id === 'HARNESS' ? '商品報價單_胸背帶.xlsx' : '商品報價單_項圈 牽繩.xlsx');
        const templateKey = `${targetSubfolder}:::${templateFileName}`;

        if (activeWorkbooks.has(templateKey)) {
          return activeWorkbooks.get(templateKey);
        }

        let b64 = null;
        if (profile.is_builtin) {
          b64 = (profile.id === 'HARNESS') ? window.CoupangTemplates?.HARNESS : window.CoupangTemplates?.LEASH;
        }
        if (!b64) {
          b64 = await window.StorageUtils.getTemplateData(profile.id);
        }
        if (!b64 && profile.template_type) {
          b64 = await window.StorageUtils.getTemplateData(profile.template_type);
        }
        // Safe fallback: if custom template binary is missing, fall back to built-in template
        if (!b64) {
          b64 = window.CoupangTemplates?.[profile.template_type] || window.CoupangTemplates?.[profile.id] || window.CoupangTemplates?.HARNESS || window.CoupangTemplates?.LEASH;
        }

        let wbInstance = null;
        let wsInstance = null;
        if (b64) {
          wbInstance = await XlsxPopulate.fromDataAsync(b64ToArrayBuffer(b64));
          wsInstance = wbInstance.sheets().find(s => s.name().startsWith('QF_')) || wbInstance.sheet(0);
        } else {
          wbInstance = await XlsxPopulate.fromBlankAsync();
          wsInstance = wbInstance.sheet(0).name(`QF_${targetSubfolder || profile.name || 'Output'}`);
        }

        const hMap = buildHeaderMap(wsInstance, 5);
        const fileName = templateFileName.replace(/\.xlsx$/i, '_auto_generate.xlsx');

        const item = {
          templateKey: templateKey,
          profile: profile,
          wb: wbInstance,
          ws: wsInstance,
          headerMap: hMap,
          nextRowIdx: 9,
          targetSubfolder: targetSubfolder,
          fileName: fileName,
          count: 0
        };

        activeWorkbooks.set(templateKey, item);
        return item;
      }

      let totalCount = 0;
      let matchCount = 0;
      let missCount = 0;
      let missBackLabelCount = 0;
      let unmatchedTemplateCount = 0;
      const typeCounts = {};

      const imageCounters = {};
      const styleImageCache = new Map(); // Map<subfolder:::baseProdName, { mainImgName, sc1ImgName, sc2ImgName, chartImgName }>

      const resolvedHeaderIndices = {
        skuIdx: null,
        collIdx: null,
        typeIdx: null,
        zhColorIdx: null,
        colorIdx: null,
        sizeIdx: null
      };

      // 1. 優先使用 currentSourceConfig 設定之全域動態欄位
      if (currentSourceConfig.collection_column) {
        resolvedHeaderIndices.collIdx = findHeaderColIdx(headers, currentSourceConfig.collection_column);
      }
      if (currentSourceConfig.type_column) {
        resolvedHeaderIndices.typeIdx = findHeaderColIdx(headers, currentSourceConfig.type_column);
      }
      if (currentSourceConfig.color_column) {
        resolvedHeaderIndices.zhColorIdx = findHeaderColIdx(headers, currentSourceConfig.color_column);
      }
      if (currentSourceConfig.size_column) {
        resolvedHeaderIndices.sizeIdx = findHeaderColIdx(headers, currentSourceConfig.size_column);
      }

      // 2. 其次檢視各 profile 之 dynamic mapping (向前相容)
      for (const p of templateProfiles) {
        const dynSku = p.field_mappings?.dynamic?.['商品條碼'] || p.field_mappings?.dynamic?.['條碼'] || p.field_mappings?.dynamic?.['EAN'] || p.field_mappings?.dynamic?.['SKU'];
        if (dynSku && !resolvedHeaderIndices.skuIdx) resolvedHeaderIndices.skuIdx = findHeaderColIdx(headers, dynSku);

        const dynColl = p.field_mappings?.dynamic?.['系列'];
        if (dynColl && !resolvedHeaderIndices.collIdx) resolvedHeaderIndices.collIdx = findHeaderColIdx(headers, dynColl);

        const dynType = p.field_mappings?.dynamic?.['種類'] || p.field_mappings?.dynamic?.['品類'] || p.field_mappings?.dynamic?.['TYPE'];
        if (dynType && !resolvedHeaderIndices.typeIdx) resolvedHeaderIndices.typeIdx = findHeaderColIdx(headers, dynType);

        const dynZhCol = p.field_mappings?.dynamic?.['中文顏色'] || p.field_mappings?.dynamic?.['顏色'];
        if (dynZhCol && !resolvedHeaderIndices.zhColorIdx) resolvedHeaderIndices.zhColorIdx = findHeaderColIdx(headers, dynZhCol);

        const dynColor = p.field_mappings?.dynamic?.['Color'] || p.field_mappings?.dynamic?.['COLOR'];
        if (dynColor && !resolvedHeaderIndices.colorIdx) resolvedHeaderIndices.colorIdx = findHeaderColIdx(headers, dynColor);

        const dynSize = p.field_mappings?.dynamic?.['尺寸'] || p.field_mappings?.dynamic?.['SIZE'];
        if (dynSize && !resolvedHeaderIndices.sizeIdx) resolvedHeaderIndices.sizeIdx = findHeaderColIdx(headers, dynSize);
      }

      // 3. 別名候選探測回退 (Fallback candidate search)
      if (!resolvedHeaderIndices.skuIdx) {
        const skuCandidates = ['SKU', '商品條碼', '條碼', 'EAN', 'EanCode', 'ean 數字', 'EAN Code', 'EAN_Code', '條碼編號', '國際條碼', 'Barcode', 'BARCODE', 'UPC', 'GTIN'];
        for (const cand of skuCandidates) {
          resolvedHeaderIndices.skuIdx = findHeaderColIdx(headers, cand);
          if (resolvedHeaderIndices.skuIdx) break;
        }
      }

      if (!resolvedHeaderIndices.collIdx) {
        const collCandidates = ['COLLECTION', '系列', 'Collection', '產品系列', '系列名稱', 'Model', 'Model_'];
        for (const cand of collCandidates) {
          resolvedHeaderIndices.collIdx = findHeaderColIdx(headers, cand);
          if (resolvedHeaderIndices.collIdx) break;
        }
      }

      if (!resolvedHeaderIndices.typeIdx) {
        const typeCandidates = ['TYPE', '種類', '品類', 'Type', '商品種類', '類別', 'Category'];
        for (const cand of typeCandidates) {
          resolvedHeaderIndices.typeIdx = findHeaderColIdx(headers, cand);
          if (resolvedHeaderIndices.typeIdx) break;
        }
      }

      if (!resolvedHeaderIndices.zhColorIdx) {
        resolvedHeaderIndices.zhColorIdx = findHeaderColIdx(headers, '中文顏色') || findHeaderColIdx(headers, '顏色') || findHeaderColIdx(headers, '顏色(中)');
      }

      if (!resolvedHeaderIndices.colorIdx) {
        resolvedHeaderIndices.colorIdx = findHeaderColIdx(headers, 'Color') || findHeaderColIdx(headers, 'COLOR') || findHeaderColIdx(headers, 'Colur') || findHeaderColIdx(headers, 'Colour') || findHeaderColIdx(headers, 'Colour_') || findHeaderColIdx(headers, '顏色(英)');
      }

      if (!resolvedHeaderIndices.sizeIdx) {
        const sizeCandidates = ['SIZE', '尺寸', 'Size', 'Size_', '規格', '尺碼'];
        for (const cand of sizeCandidates) {
          resolvedHeaderIndices.sizeIdx = findHeaderColIdx(headers, cand);
          if (resolvedHeaderIndices.sizeIdx) break;
        }
      }

      const globalFixedBrand = (currentSourceConfig.brand_fixed || '').trim();
      const globalFixedMfr = (currentSourceConfig.manufacturer_fixed || '').trim();

      for (let i = 0; i < validRowIndices.length; i++) {
        const r = validRowIndices[i];
        const currentItemNum = i + 1;

        const zhName = nameColIdx ? (getCellValue(ws.cell(r, nameColIdx)) || '').toString().trim() : '';
        if (!zhName) continue;

        const skuVal = resolvedHeaderIndices.skuIdx ? getCellValue(ws.cell(r, resolvedHeaderIndices.skuIdx)) : '';
        const skuStr = formatBarcode(skuVal);
        const rawCollection = resolvedHeaderIndices.collIdx ? (getCellValue(ws.cell(r, resolvedHeaderIndices.collIdx)) || '').toString().trim() : '';
        const rawType = resolvedHeaderIndices.typeIdx ? (getCellValue(ws.cell(r, resolvedHeaderIndices.typeIdx)) || '').toString().trim() : '';
        const rawZhColor = resolvedHeaderIndices.zhColorIdx ? (getCellValue(ws.cell(r, resolvedHeaderIndices.zhColorIdx)) || '').toString().trim() : '';
        const rawColor = resolvedHeaderIndices.colorIdx ? (getCellValue(ws.cell(r, resolvedHeaderIndices.colorIdx)) || '').toString().trim() : '';
        const rawSize = resolvedHeaderIndices.sizeIdx ? (getCellValue(ws.cell(r, resolvedHeaderIndices.sizeIdx)) || '').toString().trim() : '';
        
        const rawHints = { brand: globalFixedBrand, manufacturer: globalFixedMfr, collection: rawCollection, type: rawType, color: rawColor, zhColor: rawZhColor, size: rawSize, sku: skuStr };
        const parsed = processor.parseChineseName(zhName, rawSize, rawHints);
        let targetInfo = processor.getTargetTemplateAndCategory(zhName, rawType);

        if (targetInfo.unmatched) {
          unmatchedTemplateCount++;
          if (statUnmatchedTemplates) statUnmatchedTemplates.textContent = unmatchedTemplateCount;
          logMessage(`[略過] 商品「${zhName}」未命中任何品類報價單模板規則，已略過不寫入檔案！`, 'warning');
          continue;
        }

        totalCount++;

        let imgs = processor.getLocalImagesForProduct(parsed, rawHints);

        const currentProfile = targetInfo.profile || templateProfiles.find(p => p.id === targetInfo.template_id) || templateProfiles[0];
        const wbItem = await getOrInitWorkbook(currentProfile, targetInfo.category, targetInfo.target_subfolder);
        wbItem.count++;
        typeCounts[currentProfile.name] = (typeCounts[currentProfile.name] || 0) + 1;

        const targetSubFolder = wbItem.targetSubfolder;
        if (!filesToExport[targetSubFolder]) {
          filesToExport[targetSubFolder] = {
            images: new Map(),
            back_labels: new Map(),
            excelFiles: new Map()
          };
        }

        if (!imageCounters[targetSubFolder]) imageCounters[targetSubFolder] = 1;

        const baseProdName = getBaseProductName(zhName, rawSize);
        const styleKey = `${targetSubFolder}:::${baseProdName || zhName}`;

        let mainImgName = '';
        let sc1ImgName = '';
        let sc2ImgName = '';
        let chartImgName = '';

        if (styleImageCache.has(styleKey)) {
          const cached = styleImageCache.get(styleKey);
          mainImgName = cached.mainImgName || '';
          sc1ImgName = cached.sc1ImgName || '';
          sc2ImgName = cached.sc2ImgName || '';
          chartImgName = cached.chartImgName || '';
        } else {
          const cleanName = sanitizeFilename(baseProdName || zhName, `SKU_${skuStr || i + 1}`);

          if (imgs.main) {
            try {
              const blob = await window.ImageUtils.resizeAndPad(await window.ImageUtils.loadImage(imgs.main), 1000, 1000, '#FFFFFF');
              mainImgName = `${cleanName}主圖.jpg`;
              filesToExport[targetSubFolder].images.set(mainImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 主圖處理失敗: ${imgs.main.name}`, 'warning');
            }
          }

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

          if (imgs.chart) {
            try {
              const blob = await window.ImageUtils.ensureMinShortEdge(await window.ImageUtils.loadImage(imgs.chart));
              chartImgName = `尺寸規格表_${cleanName}.jpg`;
              filesToExport[targetSubFolder].images.set(chartImgName, blob);
            } catch(imgErr) {
              logMessage(`[警告] 尺寸圖處理失敗: ${imgs.chart.name}`, 'warning');
            }
          }

          // 恪守「寧可空白，也不要抓錯」原則：若未找到專屬尺寸圖，維持留空，絕不以主圖或情境圖冒充尺寸表
          styleImageCache.set(styleKey, {
            mainImgName,
            sc1ImgName,
            sc2ImgName,
            chartImgName
          });
        }

        let labelImgName = '';
        if (imgs.label) {
          try {
            const blob = await window.ImageUtils.ensureMinShortEdge(await window.ImageUtils.loadImage(imgs.label));
            const szSuffix = rawSize ? `_${String(rawSize).trim()}` : '';
            const cleanLabelName = sanitizeFilename(`背標_${zhName}${szSuffix}`, `背標_SKU_${skuStr || i + 1}`);
            labelImgName = `${cleanLabelName}.jpg`;
            filesToExport[targetSubFolder].back_labels.set(labelImgName, blob);
          } catch(imgErr) {
            logMessage(`[警告] 背標處理失敗: ${imgs.label.name}`, 'warning');
          }
        } else {
          missBackLabelCount++;
          logMessage(`[警告] 找不到中文背標圖: SKU=${skuStr}, 品名=${zhName}`, 'warning');
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
        const profileMappings = currentProfile.field_mappings || {};

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
          }
        };

        // Write category
        const catValue = targetInfo.category || currentProfile.category_name || '';
        setVal('細分商品種類', catValue);

        // 1. Write Global Fixed mappings from currentSourceConfig
        for (const [tKey, fVal] of Object.entries(currentSourceConfig.fixed || {})) {
          if (['細分商品種類', '品牌', '製造廠商', '系列', '顏色', '尺寸', '商品名稱'].includes(tKey)) continue;
          setVal(tKey, fVal);
        }

        // 2. Write Fixed mappings from profile (若有設定以模板設定優先覆寫)
        for (const [tKey, fVal] of Object.entries(profileMappings.fixed || {})) {
          if (['細分商品種類', '品牌', '製造廠商', '系列', '顏色', '尺寸', '商品名稱'].includes(tKey)) continue;
          if (fVal !== undefined && fVal !== null && fVal !== '') {
            setVal(tKey, fVal);
          }
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
            } else if (tKey.includes('價') || sCol.includes('價')) {
              writeVal = window.SharedUtils ? window.SharedUtils.cleanPrice(rawVal) : rawVal;
            } else if (tKey.includes('包裝尺寸')) {
              writeVal = window.SharedUtils ? window.SharedUtils.cleanPackagingDimension(rawVal) : rawVal;
            } else if (tKey.includes('包裝重量')) {
              writeVal = window.SharedUtils ? window.SharedUtils.cleanPackagingWeight(rawVal) : rawVal;
            } else {
              writeVal = (typeof rawVal === 'number') ? rawVal : (rawVal || '').toString().trim();
            }

            setVal(tKey, writeVal);
          }
        }

        // 直接由 Excel 獨立欄位與全域來源表 / Profile 設定寫入屬性（恪守「寧可空白，也不要抓錯」原則）
        const profileFixedBrand = (profileMappings.fixed?.['品牌'] || '').trim();
        const finalBrand = globalFixedBrand || profileFixedBrand || '';

        const profileFixedMfr = (profileMappings.fixed?.['製造廠商'] || '').trim();
        const finalManufacturer = globalFixedMfr || profileFixedMfr || finalBrand || '';

        const profileFixedColl = (profileMappings.fixed?.['系列'] || '').trim();
        const finalCollection = profileFixedColl || rawCollection || '';

        const finalColor = rawZhColor || rawColor || '';
        const finalSize = rawSize || '';

        if (finalBrand) setVal('品牌', finalBrand);
        if (finalManufacturer) setVal('製造廠商', finalManufacturer);
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
          if (statMissingBackLabels) statMissingBackLabels.textContent = missBackLabelCount;
          if (statUnmatchedTemplates) statUnmatchedTemplates.textContent = unmatchedTemplateCount;
          await new Promise(resolve => setTimeout(resolve, 0));
        }
      }

      // Generate binary buffers for all active workbooks
      setProgress(95, '正在壓縮並修正 Excel 格式...');
      for (const [tmplKey, item] of activeWorkbooks.entries()) {
        if (item.count > 0) {
          highlightMissingRequiredCells(item.ws, 9, item.nextRowIdx, 6);
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
      if (statMissingBackLabels) statMissingBackLabels.textContent = missBackLabelCount;
      if (statUnmatchedTemplates) statUnmatchedTemplates.textContent = unmatchedTemplateCount;

      if (totalCount === 0) {
        btnSaveToFolder.disabled = true;
        btnDownloadZip.disabled = true;
        logMessage('處理完成，但未找到任何符合篩選條件且成功配對模板的商品資料。', 'warning');
      } else {
        btnSaveToFolder.disabled = false;
        btnDownloadZip.disabled = false;
        const typesDesc = [];
        for (const [name, count] of Object.entries(typeCounts)) {
          typesDesc.push(`【${name}】: ${count} 筆`);
        }
        logMessage(`全部處理完成！共成功產生 ${totalCount} 筆 SKU（${typesDesc.join('，')}），配對成功 ${matchCount} 筆，缺圖 ${missCount} 筆${unmatchedTemplateCount > 0 ? `，未配對報價單略過 ${unmatchedTemplateCount} 筆` : ''}`, 'success');
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
    
    // 範例 1: 冷門/特殊品類 - 玩具與球類 (單一尺寸 / 特殊規格)
    const gelBall = photoFolder.folder("GEL BALL WITH ROPE 綠");
    gelBall.file("主圖.jpg", "（請放入 1000x1000 正方形商品正面主圖）");
    gelBall.file("情境圖1.jpg", "（請放入情境圖 1）");
    gelBall.file("情境圖2.jpg", "（請放入情境圖 2）");
    gelBall.file("size chart_GEL BALL WITH ROPE.png", "（請放入尺寸規格圖）");
    gelBall.file("芬蘭Rukka_附繩款彈力訓練球_綠_one_size_綠_ONE.png", "（請放入單一尺寸背標照片）");

    const ringToy = photoFolder.folder("RING TOY 暖陽黃");
    ringToy.file("主圖.jpg", "（請放入拉扯圓圈玩具主圖）");
    ringToy.file("情境圖1.jpg", "（請放入情境圖 1）");
    ringToy.file("size chart_RING TOY.png", "（請放入尺寸規格圖）");
    ringToy.file("芬蘭_Rukka_彈力拉扯圓圈玩具_暖陽黃_one_size_暖陽黃_ONE.png", "（請放入背標照片）");

    const calmDuck = photoFolder.folder("CALM DUCK TOY 白");
    calmDuck.file("主圖.jpg", "（請放入鴨子安撫玩偶主圖）");
    calmDuck.file("情境圖1.jpg", "（請放入情境圖 1）");
    calmDuck.file("size chart_CALM DUCK TOY.png", "（請放入尺寸規格圖）");
    calmDuck.file("芬蘭_Rukka_鴨子安撫玩偶_白_one_size_白_ONE.png", "（請放入背標照片）");

    // 範例 2: 服飾與雨衣 (多尺寸、紅黑顏色對比)
    const raincoatBlack = photoFolder.folder("HAYTON X RAINCOAT 黑");
    raincoatBlack.file("主圖.jpg", "（請放入黑色雨衣主圖）");
    raincoatBlack.file("情境圖1.jpg", "（請放入情境圖 1）");
    raincoatBlack.file("size chart_haytonXraincoat.png", "（請放入尺寸規格圖）");
    raincoatBlack.file("芬蘭_Rukka_HAYTON_X_輕量雨衣_黑_25_黑_25_0.png", "（請放入 25 號背標照片）");
    raincoatBlack.file("芬蘭_Rukka_HAYTON_X_輕量雨衣_黑_30_黑_30.png", "（請放入 30 號背標照片）");

    // 範例 3: 生活與訓練配件 (吸水浴袍 XXS~XXL 多尺寸)
    const bathrobeGreen = photoFolder.folder("AALTO PET BATHROBE 蘋果綠");
    bathrobeGreen.file("主圖.jpg", "（請放入蘋果綠浴袍主圖）");
    bathrobeGreen.file("情境圖1.jpg", "（請放入情境圖 1）");
    bathrobeGreen.file("size chart_AALTO PET BATHROBE.png", "（請放入尺寸規格圖）");
    bathrobeGreen.file("芬蘭_Rukka_吸水_浴袍_蘋果綠_M_蘋果綠_M.png", "（請放入 M 號背標照片）");
    bathrobeGreen.file("芬蘭_Rukka_吸水_浴袍_蘋果綠_L_蘋果綠_L.png", "（請放入 L 號背標照片）");

    const content = await zip.generateAsync({ type: "blob" });
    saveAs(content, "Coupang_Photo_Sample_Structure.zip");
    logMessage("範例資料夾結構已下載！", "success");
  });
});
