class MyFamilyProcessor {
  constructor(photoFilesArray, excelFileBuffer, colorAliases, zhColorMap, targetCombos, collectionAliases = {}) {
    this.photoFiles = photoFilesArray; 
    this.excelBuffer = excelFileBuffer;
    this.colorAliases = colorAliases;
    this.zhColorMap = zhColorMap;
    this.targetCombos = targetCombos;
    this.collectionAliases = collectionAliases;
    
    this.folderMap = new Map();
    for (const file of this.photoFiles) {
      const pathParts = file.webkitRelativePath.split('/');
      pathParts.pop(); 
      const dirPath = pathParts.join('/');
      
      if (!this.folderMap.has(dirPath)) {
        this.folderMap.set(dirPath, []);
      }
      this.folderMap.get(dirPath).push(file);
    }
  }

  normStr(s) {
    if (!s) return '';
    s = s.toUpperCase();
    s = s.replace(/[\（\(\）\)]/g, '');
    return s.replace(/ /g, '').replace(/_/g, '').replace(/-/g, '');
  }

  getTargetTemplateAndCategory(prodName, typeStr) {
    const typeUpper = (typeStr || '').toUpperCase();
    const nameUpper = (prodName || '').toUpperCase();

    if (typeUpper.includes('HARNESS') || prodName.includes('胸背帶') || nameUpper.includes('HARNESS')) {
      return {
        template_file: '商品報價單_胸背帶.xlsx',
        output_file: '商品報價單_胸背帶_auto_generate.xlsx',
        category: '寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)',
        target_subfolder: '胸背帶'
      };
    } else if (typeUpper.includes('LEASH') || prodName.includes('牽繩') || nameUpper.includes('LEASH')) {
      return {
        template_file: '商品報價單_項圈 牽繩.xlsx',
        output_file: '商品報價單_項圈 牽繩_auto_generate.xlsx',
        category: '寵物用品>犬貓通用>項圈/伸縮牽繩>牽繩 (66027)',
        target_subfolder: '牽繩'
      };
    } else {
      return {
        template_file: '商品報價單_項圈 牽繩.xlsx',
        output_file: '商品報價單_項圈 牽繩_auto_generate.xlsx',
        category: '寵物用品>犬貓通用>項圈/伸縮牽繩>項圈 (66025)',
        target_subfolder: '牽繩'
      };
    }
  }

  extractColorFromZhName(zhName, prodSize = "") {
    if (!zhName) return "";
    let s = String(zhName).trim();
    let parts = s.split(/\s+/);
    
    if (parts.length >= 2) {
      let last = parts[parts.length - 1];
      let szStr = String(prodSize || '').trim().toUpperCase();
      let cand;
      
      let isSize = (szStr && last.toUpperCase() === szStr) || 
                   /^[0-9]*X*[SML]+$/i.test(last) || 
                   /([0-9]+(?:cm)?\/[0-9.]+(?:mm|cm)?|[0-9]+-[0-9.]+|[0-9]+\/[0-9]+mm)/i.test(last);
                   
      if (isSize && parts.length > 2) {
        cand = parts[parts.length - 2];
      } else if (isSize) {
        cand = parts[0];
      } else {
        cand = last;
      }
      
      cand = cand.trim();
      if (szStr && cand.toUpperCase().endsWith(szStr)) {
        cand = cand.slice(0, -szStr.length).trim();
      }
      cand = cand.replace(/([0-9]*X*[SML]+|[0-9]+(?:cm)?\/[0-9.]+(?:mm|cm)?|[0-9]+-[0-9.]+|[0-9]+\/[0-9]+mm)$/i, '').trim();
      return cand;
    }
    return "";
  }

  
  parseChineseName(prodName, excelSize = "") {
    const isSizeStr = (s) => {
      if (!s) return false;
      const sClean = s.replace(/Ｍ/g, 'M').replace(/Ｓ/g, 'S').replace(/Ｌ/g, 'L').replace(/Ｘ/g, 'X').trim().toUpperCase();
      if (/^[0-9]*X*[SML]+$/i.test(sClean)) return true;
      if (/([0-9]+(?:cm)?\/[0-9.]+(?:mm|cm)?|[0-9]+-[0-9.]+|[0-9]+\/[0-9]+mm|[0-9]+_[0-9.]+|UNISIZE)/i.test(sClean)) return true;
      return false;
    };

    const parts = (prodName || '').trim().split(/[\s]+/).filter(p => p.length > 0);
    if (parts.length < 3) {
      return { success: false, brand: '', collection: '', type: '', color: '', size: '' };
    }

    let size = '';
    const last = parts[parts.length - 1];
    if (isSizeStr(last) || (excelSize && last.toUpperCase() === excelSize.trim().toUpperCase())) {
      size = parts.pop();
    } else if (excelSize) {
      size = excelSize.trim();
    }

    let color = parts.pop() || '';
    let type = '';
    let collection = '';

    // 處理顏色與品類黏在一起的情況 (例: 頂級皮革狗項圈紫桃紅 或 貓用防開安全項圈)
    if (color && /項圈|胸背帶|牽繩|背帶/.test(color)) {
      const splitMatch = color.match(/^(.*?(?:狗項圈|貓項圈|貓咪項圈|胸背帶|項圈|牽繩|背帶))(.*)$/);
      if (splitMatch && splitMatch[2].trim()) {
        type = splitMatch[1].trim();
        color = splitMatch[2].trim();
        collection = parts.pop() || '';
      } else {
        type = color;
        color = '';
        collection = parts.pop() || '';
      }
    } else {
      type = parts.pop() || '';
      collection = parts.pop() || '';
    }

    if (collection) {
      collection = collection.replace(/系列$/g, '').trim();
    }
    const brand = parts.join(' ');

    return { success: true, brand, collection, type, color, size };
  }

  getLocalImagesForProduct(parsedData, rawHints = {}) {
    const collZh = (parsedData.collection || '').replace(/系列$/g, '').trim();
    const collEn = (rawHints.collection || '').trim();
    const typZh = (parsedData.type || '').trim();
    const typEn = (rawHints.type || '').trim();
    const typeStr = (typZh + ' ' + typEn).toUpperCase();
    const colorZh = (parsedData.color || '').trim();
    const colorEn = (rawHints.color || '').trim();
    const sizeStr = (parsedData.size || rawHints.size || '').trim().toUpperCase();

    // 0. Collect all collection aliases from UI config
    let collectionAliases = [collZh, collEn].filter(Boolean);
    for (const [canonical, aliasList] of Object.entries(this.collectionAliases || {})) {
      const match = aliasList.some(a => {
        const cleanA = a.replace(/系列$/g, '').trim().toUpperCase();
        return (
          (collZh && cleanA === collZh.toUpperCase()) ||
          (collEn && cleanA === collEn.toUpperCase()) ||
          (collZh && collZh.includes(cleanA)) ||
          (collZh && cleanA.includes(collZh))
        );
      });
      if (match) {
        collectionAliases.push(canonical, ...aliasList.map(s => s.replace(/系列$/g, '').trim()));
      }
    }
    collectionAliases = [...new Set(collectionAliases.map(s => s.toUpperCase().trim()).filter(Boolean))];

    // 1. Collect all color aliases from UI config
    let colorAliases = [colorZh, colorEn].filter(Boolean);
    for (const [canonical, aliasList] of Object.entries(this.colorAliases || {})) {
      const match = aliasList.some(a => 
        (colorZh && a.toUpperCase() === colorZh.toUpperCase()) ||
        (colorEn && a.toUpperCase() === colorEn.toUpperCase()) ||
        (colorZh && colorZh.includes(a))
      );
      if (match) {
        colorAliases.push(canonical, ...aliasList);
      }
    }
    colorAliases = [...new Set(colorAliases.map(s => s.toUpperCase().trim()).filter(Boolean))];

    let candidates = [];

    for (const [dirPath, files] of this.folderMap.entries()) {
      const relU = dirPath.toUpperCase();
      let score = 0;

      // Type matching
      if (typeStr.includes('胸背帶') || typeStr.includes('HARNESS')) {
        if (relU.includes('HARNESS')) {
          score += 30;
          if ((typeStr.includes('X') || typeStr.includes('Ｘ')) && relU.includes('X HARNESS')) score += 20;
          if ((typeStr.includes('H') || typeStr.includes('Ｈ')) && relU.includes('H HARNESS')) score += 20;
        }
      } else if (typeStr.includes('牽繩') || typeStr.includes('LEASH')) {
        if (relU.includes('LEASH')) score += 30;
      } else if (typeStr.includes('貓') || typeStr.includes('CAT')) {
        if (relU.includes('CAT') && relU.includes('COLLAR')) score += 30;
      } else {
        if (relU.includes('COLLAR') && !relU.includes('CAT')) score += 30;
      }

      if (score === 0) continue;

      // Collection matching
      let collMatched = false;
      for (const cal of collectionAliases) {
        if (relU.includes(cal)) {
          score += 40;
          collMatched = true;
          break;
        }
      }
      if (!collMatched) {
        for (const cal of collectionAliases) {
          if (files.some(f => f.name.toUpperCase().includes(cal))) {
            score += 35;
            collMatched = true;
            break;
          }
        }
      }

      // Color matching
      let colorMatched = false;
      if (colorEn && relU.includes(colorEn.toUpperCase())) {
        score += 40;
        colorMatched = true;
      }
      for (const al of colorAliases) {
        if (relU.includes(al)) {
          score += 35;
          colorMatched = true;
          break;
        }
      }
      if (!colorMatched && colorZh && files.some(f => f.name.includes(colorZh))) {
        score += 30;
        colorMatched = true;
      }

      // Size matching in files (back labels)
      if (sizeStr && files.some(f => f.name.toUpperCase().includes(sizeStr))) {
        score += 15;
      }

      if (collMatched || colorMatched || score >= 50) {
        candidates.push({ score, dirPath, files });
      }
    }

    let matchedFolder = null;
    let matchedFiles = [];

    if (candidates.length > 0) {
      candidates.sort((a, b) => b.score - a.score);
      matchedFolder = candidates[0].dirPath;
      matchedFiles = candidates[0].files;
    }

    if (!matchedFolder) {
      return { main: null, sc1: null, sc2: null, label: null };
    }

    let mainImg = null;
    let sc1 = null;
    let sc2 = null;
    let chart = null;
    let label = null;

    for (const f of matchedFiles) {
      const fn = f.name;
      const fnL = fn.toLowerCase();
      
      if (fn.includes('主圖') || fnL.startsWith('main') || fn.includes('主圖1')) {
        if (!mainImg) mainImg = f;
      } else if (fn.includes('情境圖 1') || fn.includes('情境圖1') || fn.includes('情境1') || fn === '情境圖.jpg' || fn === '情境圖.png') {
        if (!sc1) sc1 = f;
      } else if (fn.includes('情境圖 2') || fn.includes('情境圖2') || fn.includes('情境2')) {
        if (!sc2) sc2 = f;
      } else if (fn.includes('尺寸圖') || fn.includes('尺寸表') || fnL.includes('chart') || fn.includes('尺寸規格')) {
        if (!chart) chart = f;
      }
    }

    // 精確匹配該尺寸專屬的背標圖檔 (嚴格排除尺寸代碼子字串誤判，如 XS 誤匹配 S/M 等)
    const normTargetSize = sizeStr ? sizeStr.replace(/Ｍ/g, 'M').replace(/Ｓ/g, 'S').replace(/Ｌ/g, 'L').replace(/Ｘ/g, 'X').replace(/號/g, '').trim().toUpperCase() : '';
    
    if (normTargetSize) {
      const backLabelCandidates = matchedFiles.filter(f => {
        const fn = f.name;
        const fl = fn.toLowerCase();
        if (fn.includes('主圖') || fl.startsWith('main') || fn.includes('主圖1')) return false;
        if (fn.includes('情境') || fl.includes('scene')) return false;
        if (fn.includes('尺寸圖') || fn.includes('尺寸表') || fl.includes('chart') || fn.includes('尺寸規格')) return false;
        if (fn === '.DS_Store' || fl.endsWith('.db')) return false;
        return true;
      });

      let bestScore = -1;
      for (const f of backLabelCandidates) {
        const fn = f.name;
        const nameNoExt = fn.substring(0, fn.lastIndexOf('.')).toUpperCase();
        let score = 0;

        if (normTargetSize === 'UNISIZE' || normTargetSize === '單一尺寸') {
          if (nameNoExt.includes('單一尺寸') || nameNoExt.includes('UNISIZE')) {
            score = 100;
          }
        } else {
          // 尺寸規格格式 (如 110-1.2, 110/1.6, 60cm_28mm)
          const cleanTargetDim = normTargetSize.replace(/CM|MM/g, '').replace(/[\s\/\-_*xX\.]+/g, '_');
          const cleanNameDim = nameNoExt.replace(/CM|MM/g, '').replace(/[\s\/\-_*xX\.]+/g, '_');
          if (cleanTargetDim.length >= 3 && cleanNameDim.includes(cleanTargetDim)) {
            score = 90;
          }

          // 服飾標準尺寸邊界匹配 (4XS, 3XS, 2XS, XS, S, SM, M, ML, L, XL, 2XL, 3XL, 4XL)
          const escaped = normTargetSize.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
          const regexExact = new RegExp(`(?:^|[\\s_\\-\\(\\[])${escaped}(?:$|[\\s_\\-\\)\\]]|_?背標)`, 'i');
          if (regexExact.test(nameNoExt)) {
            score = 100;
          }
        }

        if (score > bestScore) {
          bestScore = score;
          label = f;
        }
      }
    }

    // 若當前顏色資料夾中沒有尺寸圖，向上層/同系列資料夾遞迴尋找尺寸表
    if (!chart && matchedFolder) {
      const parentPath = matchedFolder.substring(0, matchedFolder.lastIndexOf('/'));
      for (const [dPath, files] of this.folderMap.entries()) {
        if (dPath.startsWith(parentPath) || (collZh && dPath.toUpperCase().includes(collZh.toUpperCase())) || (collEn && dPath.toUpperCase().includes(collEn.toUpperCase()))) {
          for (const f of files) {
            const fn = f.name;
            const fnL = fn.toLowerCase();
            if (fn.includes('尺寸圖') || fn.includes('尺寸表') || fnL.includes('chart') || fn.includes('尺寸規格')) {
              chart = f;
              break;
            }
          }
        }
        if (chart) break;
      }
    }

    if (!mainImg) {
      const jpgs = matchedFiles.filter(f => {
        const name = f.name.toLowerCase();
        return (name.endsWith('.jpg') || name.endsWith('.jpeg') || name.endsWith('.png')) && !name.includes('情境') && !name.includes('尺寸') && !name.includes('背標');
      });
      if (jpgs.length > 0) mainImg = jpgs[0];
    }

    if (!sc1) {
      const extraJpgs = matchedFiles.filter(f => {
        const name = f.name.toLowerCase();
        return (name.endsWith('.jpg') || name.endsWith('.jpeg') || name.endsWith('.png')) && f !== mainImg && !name.includes('尺寸') && !name.includes('背標');
      });
      if (extraJpgs.length > 0) {
        sc1 = extraJpgs[0];
        if (extraJpgs.length > 1 && !sc2) {
          sc2 = extraJpgs[1];
        }
      }
    }

    return { main: mainImg, sc1: sc1, sc2: sc2, chart: chart, label: label };
  }

  static isCategoryOrNonColor(folderName) {
    if (!folderName) return true;
    const fUpper = folderName.toUpperCase().trim();
    const badKeywords = [
      '圖片', 'PHOTO', 'IMAGE', 'IMAGES', 'BACK_LABEL', '背標', '尺寸', 'CHART', 'TABLE',
      'HARNESS', 'COLLAR', 'LEASH', '項圈', '胸背帶', '牽繩', '背帶',
      'NYLON', 'LEATHER', 'WITH ROPE', 'WITH CHAIN', 'ROPE LEASH', '平板LEASH',
      '已成功上架', 'DOG', 'CAT'
    ];
    for (const kw of badKeywords) {
      if (fUpper.includes(kw)) return true;
    }
    if (/^[0-9._\s-]+$/.test(fUpper)) return true;
    return false;
  }

  static scanFolderStructure(photoFilesArray) {
    const collections = new Set();
    const colors = new Set();

    for (const file of photoFilesArray) {
      const relPath = file.webkitRelativePath || '';
      if (!relPath) continue;

      const parts = relPath.split('/').filter(p => p.trim().length > 0);
      if (parts.length <= 1) continue;

      // 移除檔名，保留資料夾層級
      parts.pop();

      // 若第一層是 Photo 或 photo，略過 root folder
      let dirParts = [...parts];
      if (dirParts.length > 0 && dirParts[0].toUpperCase() === 'PHOTO') {
        dirParts.shift();
      }

      if (dirParts.length === 0) continue;

      // 第一層視為系列 (Collection)
      const topDir = dirParts[0].trim();
      if (topDir && !this.isCategoryOrNonColor(topDir)) {
        collections.add(topDir);
      }

      // 若只有一層資料夾且已被當作系列，則不作為顏色
      if (dirParts.length === 1 && collections.has(topDir)) {
        continue;
      }

      // 末端資料夾通常為顏色 (Color) 或單層商品名
      const leafDir = dirParts[dirParts.length - 1].trim();

      // 處理特殊情況如 "FIRENZE cat collar 天空藍"
      let candColor = leafDir;
      const catCollarMatch = candColor.match(/collar\s*(.+)$/i);
      if (catCollarMatch) {
        candColor = catCollarMatch[1].trim();
      }

      // 處理規格數字如 "Blue110_1.6_110_2.2"
      const dimMatch = candColor.match(/^([A-Za-z\u4e00-\u9fa5\s&（）\(\)]+?)(?:[0-9]+.*)?$/);
      if (dimMatch && dimMatch[1].trim()) {
        candColor = dimMatch[1].trim();
      }

      // 若過濾後不是品類詞、且不是已知的系列名稱，則加入顏色
      if (candColor && !this.isCategoryOrNonColor(candColor) && !collections.has(candColor)) {
        colors.add(candColor);
      }
    }

    return {
      collections: Array.from(collections),
      colors: Array.from(colors)
    };
  }

  static extractMappingsFromExcel(workbook, headerRow = 3, rowStart = 4, filterColName = '中文背標') {
    const sheetNames = ['MYFAMILY', 'MY FAMILY', 'My Family', 'Sheet1'];
    let ws = null;
    for (const name of sheetNames) {
      ws = workbook.sheet(name);
      if (ws) break;
    }
    if (!ws) ws = workbook.sheet(0);
    if (!ws) return { collections: {}, colors: {} };

    const usedRange = ws.usedRange();
    if (!usedRange) return { collections: {}, colors: {} };
    const maxCol = usedRange.endCell().columnNumber();
    const totalRows = usedRange.endCell().rowNumber();

    const headers = {};
    for (let c = 1; c <= maxCol; c++) {
      const val = ws.cell(headerRow, c).value();
      if (val) {
        const key = val.toString().trim();
        if (!(key in headers)) headers[key] = c;
      }
    }

    // 尋找目標欄位 index
    const colNameIdx = headers['中文品名'] || headers['商品名稱'] || headers['品名'];
    const colColorIdx = headers['COLOR'] || headers['顏色'] || headers['Color'];
    const colCollIdx = headers['COLLECTION'] || headers['系列'] || headers['Collection'];
    const colSizeIdx = headers['SIZE'] || headers['尺寸'] || headers['Size'];

    const getVal = (r, colIdx) => {
      if (!colIdx) return '';
      const cell = ws.cell(r, colIdx);
      const v = cell.value();
      if (v === null || v === undefined) return '';
      if (typeof v === 'object' && typeof v.text === 'function') return v.text().trim();
      return v.toString().trim();
    };

    const collectionMap = {}; // { "HERMITAGE": Set(["隱士", ...]) }
    const colorMap = {};      // { "CAMEL": Set(["可可棕", ...]) }

    const dummyProcessor = new MyFamilyProcessor([], null, {}, [], [], {});

    for (let r = rowStart; r <= totalRows; r++) {
      const zhName = getVal(r, colNameIdx);
      if (!zhName) continue;

      const rawColor = getVal(r, colColorIdx);
      const rawColl = getVal(r, colCollIdx);
      const rawSize = getVal(r, colSizeIdx);

      const parsed = dummyProcessor.parseChineseName(zhName, rawSize);
      let zhColor = parsed.color || dummyProcessor.extractColorFromZhName(zhName, rawSize);
      const zhColl = (parsed.collection || '').replace(/系列$/g, '').trim();

      if (zhColor && MyFamilyProcessor.isCategoryOrNonColor(zhColor)) {
        zhColor = '';
      }

      if (rawColl && zhColl && !MyFamilyProcessor.isCategoryOrNonColor(rawColl)) {
        const cleanCollKey = rawColl.trim();
        if (!collectionMap[cleanCollKey]) collectionMap[cleanCollKey] = new Set();
        collectionMap[cleanCollKey].add(zhColl);
      }

      if (rawColor && zhColor && !MyFamilyProcessor.isCategoryOrNonColor(rawColor)) {
        const cleanColorKey = rawColor.trim();
        if (!colorMap[cleanColorKey]) colorMap[cleanColorKey] = new Set();
        colorMap[cleanColorKey].add(zhColor);
      }
    }

    // 轉為普通陣列物件
    const resCollections = {};
    for (const [k, vSet] of Object.entries(collectionMap)) {
      resCollections[k] = Array.from(vSet).filter(Boolean);
    }
    const resColors = {};
    for (const [k, vSet] of Object.entries(colorMap)) {
      resColors[k] = Array.from(vSet).filter(Boolean);
    }

    return {
      collections: resCollections,
      colors: resColors
    };
  }

  static mergeScannedAliases(existingCollections = {}, existingColors = {}, scannedFolder = { collections: [], colors: [] }, excelData = { collections: {}, colors: {} }) {
    const newCollections = {};
    for (const [k, v] of Object.entries(existingCollections)) {
      if (!MyFamilyProcessor.isCategoryOrNonColor(k)) {
        newCollections[k] = (v || []).filter(al => !MyFamilyProcessor.isCategoryOrNonColor(al));
      }
    }

    const newColors = {};
    for (const [k, v] of Object.entries(existingColors)) {
      if (!MyFamilyProcessor.isCategoryOrNonColor(k)) {
        newColors[k] = (v || []).filter(al => !MyFamilyProcessor.isCategoryOrNonColor(al));
      }
    }

    const norm = s => (s || '').toUpperCase().replace(/[\（\(\）\)\s_-]/g, '').replace(/系列$/g, '');

    let newCollCount = 0;
    let newColorCount = 0;
    let matchedExcelCount = 0;

    // 1. 合併掃描到的系列 (Collections)
    for (const collName of scannedFolder.collections || []) {
      if (MyFamilyProcessor.isCategoryOrNonColor(collName)) continue;
      const nColl = norm(collName);
      if (!nColl) continue;

      let foundKey = Object.keys(newCollections).find(k => norm(k) === nColl);
      if (!foundKey) {
        newCollections[collName] = [collName];
        foundKey = collName;
        newCollCount++;
      }
    }

    // 2. 合併 Excel 中發現的系列中英文對應
    for (const [rawColl, zhColls] of Object.entries(excelData.collections || {})) {
      if (MyFamilyProcessor.isCategoryOrNonColor(rawColl)) continue;
      const nColl = norm(rawColl);
      if (!nColl) continue;

      const validZhColls = (zhColls || []).filter(zh => !MyFamilyProcessor.isCategoryOrNonColor(zh));

      let foundKey = Object.keys(newCollections).find(k => norm(k) === nColl);
      if (!foundKey) {
        newCollections[rawColl] = [rawColl, ...validZhColls];
        newCollCount++;
        matchedExcelCount += validZhColls.length;
      } else {
        for (const zh of validZhColls) {
          if (zh && !newCollections[foundKey].includes(zh)) {
            newCollections[foundKey].push(zh);
            matchedExcelCount++;
          }
        }
      }
    }

    // 3. 合併掃描到的顏色 (Colors)
    for (const colorName of scannedFolder.colors || []) {
      if (MyFamilyProcessor.isCategoryOrNonColor(colorName)) continue;
      const nCol = norm(colorName);
      if (!nCol) continue;

      let foundKey = Object.keys(newColors).find(k => {
        if (norm(k) === nCol) return true;
        return (newColors[k] || []).some(al => norm(al) === nCol);
      });

      if (!foundKey) {
        newColors[colorName] = [colorName];
        newColorCount++;
      }
    }

    // 4. 合併 Excel 中發現的顏色中英文對應
    for (const [rawColor, zhColors] of Object.entries(excelData.colors || {})) {
      if (MyFamilyProcessor.isCategoryOrNonColor(rawColor)) continue;
      const nCol = norm(rawColor);
      if (!nCol) continue;

      const validZhColors = (zhColors || []).filter(zh => !MyFamilyProcessor.isCategoryOrNonColor(zh));

      let foundKey = Object.keys(newColors).find(k => {
        if (norm(k) === nCol) return true;
        return (newColors[k] || []).some(al => norm(al) === nCol);
      });

      if (!foundKey) {
        newColors[rawColor] = [rawColor, ...validZhColors];
        newColorCount++;
        matchedExcelCount += validZhColors.length;
      } else {
        for (const zh of validZhColors) {
          if (zh && !newColors[foundKey].includes(zh)) {
            newColors[foundKey].push(zh);
            matchedExcelCount++;
          }
        }
      }
    }

    return {
      collectionAliases: newCollections,
      colorAliases: newColors,
      stats: {
        newCollCount,
        newColorCount,
        matchedExcelCount
      }
    };
  }
}

window.MyFamilyProcessor = MyFamilyProcessor;
