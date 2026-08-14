class CoupangProcessor {
  constructor(photoFilesArray, excelFileBuffer, colorAliases = {}, collectionAliases = {}, templateProfiles = []) {
    // Backwards compatibility check if old signature is used
    if (Array.isArray(arguments[3]) && Array.isArray(arguments[4])) {
      collectionAliases = arguments[5] || {};
      templateProfiles = arguments[6] || [];
    }
    this.photoFiles = photoFilesArray || []; 
    this.excelBuffer = excelFileBuffer;
    this.colorAliases = colorAliases || {};
    this.collectionAliases = collectionAliases || {};
    this.templateProfiles = templateProfiles || [];
    this.categoryRules = this.templateProfiles;
    
    this.folderMap = new Map();
    for (const file of (this.photoFiles || [])) {
      const pathParts = file.webkitRelativePath.split('/');
      pathParts.pop(); 
      const dirPath = pathParts.join('/');
      
      if (!this.folderMap.has(dirPath)) {
        this.folderMap.set(dirPath, []);
      }
      this.folderMap.get(dirPath).push(file);
    }
  }

  getAllCategoryKeywords() {
    return CoupangProcessor.getAllCategoryKeywords(this.templateProfiles);
  }

  static getAllCategoryKeywords(templateProfiles = []) {
    const profiles = (templateProfiles && templateProfiles.length > 0)
      ? templateProfiles
      : (window.AppConfig ? window.AppConfig.getDefaultProfiles() : []);
      
    const baseKeywords = [
      'HARNESS', 'COLLAR', 'LEASH', '項圈', '胸背帶', '牽繩', '背帶', '拉繩',
      'CAT COLLAR', 'DOG COLLAR', 'X HARNESS', 'H HARNESS', 'ROPE LEASH', '貓項圈', '狗項圈', '貓咪項圈'
    ];
    const keywords = [...baseKeywords];
    
    for (const p of profiles) {
      if (p.name) keywords.push(p.name.toString().trim());
      if (p.keywords && Array.isArray(p.keywords)) {
        keywords.push(...p.keywords.map(k => (k || '').toString().trim()));
      }
    }
    
    const unique = [...new Set(keywords.filter(Boolean))];
    return unique.sort((a, b) => b.length - a.length);
  }

  normStr(s) {
    if (!s) return '';
    s = s.toUpperCase();
    s = s.replace(/[\（\(\）\)]/g, '');
    return s.replace(/ /g, '').replace(/_/g, '').replace(/-/g, '');
  }

  getTargetTemplateAndCategory(prodName = '', typeStr = '') {
    const typeUpper = (typeStr || '').toUpperCase().trim();
    const nameUpper = (prodName || '').toUpperCase().trim();
    const fullText = (typeUpper + ' ' + nameUpper).trim();

    const profiles = (this.templateProfiles && this.templateProfiles.length > 0)
      ? this.templateProfiles
      : (window.AppConfig ? window.AppConfig.getDefaultProfiles() : []);

    for (const p of profiles) {
      const keywords = (p.keywords || [p.name]).map(k => (k || '').toString().trim().toUpperCase()).filter(Boolean);
      const isMatch = keywords.some(k => fullText.includes(k) || typeUpper.includes(k) || nameUpper.includes(k));
      if (isMatch) {
        const tmplId = p.id || p.template_id || p.template_type || 'LEASH';
        const tmplType = (p.template_type || p.id || 'LEASH').toUpperCase();
        const tmplFile = p.template_file_name || p.template_name || (tmplType === 'HARNESS' ? '商品報價單_胸背帶.xlsx' : '商品報價單_項圈 牽繩.xlsx');
        const outputFile = tmplFile.replace(/\.xlsx$/i, '_auto_generate.xlsx');
        return {
          template_id: tmplId,
          template_type: tmplType,
          template_file: tmplFile,
          output_file: outputFile,
          category: p.category_name || (tmplType === 'HARNESS' ? '寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)' : '寵物用品>犬貓通用>項圈/伸縮牽繩>項圈 (66025)'),
          target_subfolder: p.subfolder || p.name || (tmplType === 'HARNESS' ? '胸背帶' : '項圈牽繩'),
          matched_rule: p.name || '',
          unmatched: false
        };
      }
    }

    return {
      template_id: 'UNMATCHED',
      template_type: 'LEASH',
      template_file: '商品報價單_項圈 牽繩.xlsx',
      output_file: '商品報價單_項圈 牽繩_auto_generate.xlsx',
      category: '',
      target_subfolder: '未分類品項',
      matched_rule: '',
      unmatched: true
    };
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

  parseChineseName(prodName, excelSize = "", rawHints = {}) {
    const isSizeStr = (s) => {
      if (!s) return false;
      const sClean = s.replace(/Ｍ/g, 'M').replace(/Ｓ/g, 'S').replace(/Ｌ/g, 'L').replace(/Ｘ/g, 'X').trim().toUpperCase();
      if (/^[0-9]*X*[SML]+$/i.test(sClean)) return true;
      if (/([0-9]+(?:cm)?\/[0-9.]+(?:mm|cm)?|[0-9]+-[0-9.]+|[0-9]+\/[0-9]+mm|[0-9]+_[0-9.]+|UNISIZE)/i.test(sClean)) return true;
      return false;
    };

    let brand = (rawHints && rawHints.brand) ? String(rawHints.brand).trim() : '';
    let collection = (rawHints && rawHints.collection) ? String(rawHints.collection).trim() : '';
    let type = (rawHints && rawHints.type) ? String(rawHints.type).trim() : '';
    let color = (rawHints && rawHints.color) ? String(rawHints.color).trim() : '';
    let size = (rawHints && rawHints.size) ? String(rawHints.size).trim() : (excelSize ? String(excelSize).trim() : '');

    const parts = (prodName || '').trim().split(/[\s]+/).filter(p => p.length > 0);
    
    // 若品名有 3 個以上詞段，嘗試標準位置剖析：[品牌] [系列] [品類] [顏色] [尺寸]
    if (parts.length >= 3) {
      let parsedSize = '';
      const last = parts[parts.length - 1];
      if (isSizeStr(last) || (excelSize && last.toUpperCase() === excelSize.trim().toUpperCase())) {
        parsedSize = parts.pop();
      } else if (excelSize) {
        parsedSize = excelSize.trim();
      }
      if (parsedSize) size = parsedSize;

      let parsedColor = parts.pop() || '';
      let parsedType = '';
      let parsedColl = '';

      // 動態由所有 Profile 關鍵字組合正則，處理顏色與品類黏在一起的情況 (例: 頂級皮革狗項圈紫桃紅 或 雲朵雨衣黃色)
      const catKws = this.getAllCategoryKeywords();
      const catPattern = catKws.map(k => k.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')).join('|');
      const catRegex = new RegExp(`^(.*?(?:${catPattern}))(.*)$`, 'i');

      if (parsedColor && catRegex.test(parsedColor)) {
        const splitMatch = parsedColor.match(catRegex);
        if (splitMatch && splitMatch[2].trim()) {
          parsedType = splitMatch[1].trim();
          parsedColor = splitMatch[2].trim();
          parsedColl = parts.pop() || '';
        } else {
          parsedType = parsedColor;
          parsedColor = '';
          parsedColl = parts.pop() || '';
        }
      } else {
        parsedType = parts.pop() || '';
        parsedColl = parts.pop() || '';
      }

      if (parsedColl) {
        collection = parsedColl.replace(/系列$/g, '').trim();
      }
      if (parsedType) type = parsedType;
      if (parsedColor) color = parsedColor;
      if (parts.length > 0) {
        brand = parts.join(' ');
      }

      return { success: true, brand, collection, type, color, size };
    }

    // 彈性回退 (Fallback): 當品名詞段小於 3 個（非標準結構品名），以來源 Excel 資訊與關鍵字搜尋補齊
    if (prodName) {
      const catKws = this.getAllCategoryKeywords();
      for (const kw of catKws) {
        if (prodName.toUpperCase().includes(kw.toUpperCase())) {
          if (!type) type = kw;
          break;
        }
      }
      const extractedCol = this.extractColorFromZhName(prodName, size);
      if (extractedCol && !color) color = extractedCol;
    }

    const hasInfo = Boolean(type || color || collection || brand || size);
    return { 
      success: hasInfo, 
      brand, 
      collection: (collection || '').replace(/系列$/g, '').trim(), 
      type, 
      color, 
      size 
    };
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
          (collZh && cleanA.includes(cleanA)) ||
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

    // 2. 動態取得所屬品類 Profile 之所有關鍵字與排除衝突之其他品類關鍵字
    const targetProfileInfo = this.getTargetTemplateAndCategory(parsedData.type || '', typeStr);
    const profiles = (this.templateProfiles && this.templateProfiles.length > 0)
      ? this.templateProfiles
      : (window.AppConfig ? window.AppConfig.getDefaultProfiles() : []);

    let currentProfile = profiles.find(p => p.id === targetProfileInfo.template_id);
    if (!currentProfile) {
      currentProfile = profiles.find(p => p.id === targetProfileInfo.template_type) || profiles[0];
    }

    const targetKeywords = (currentProfile?.keywords || [currentProfile?.name || ''])
      .map(k => (k || '').toString().toUpperCase().trim())
      .filter(Boolean);

    const otherProfileKeywords = [];
    for (const p of profiles) {
      if (p.id !== currentProfile?.id) {
        if (p.name) otherProfileKeywords.push(p.name.toString().toUpperCase().trim());
        if (p.keywords && Array.isArray(p.keywords)) {
          otherProfileKeywords.push(...p.keywords.map(k => (k || '').toString().toUpperCase().trim()));
        }
      }
    }

    let candidates = [];

    for (const [dirPath, files] of this.folderMap.entries()) {
      const relU = dirPath.toUpperCase();
      let score = 0;

      // Type matching: 動態判定路徑是否符合該品類關鍵字
      let typeMatched = false;
      if (targetKeywords.length > 0) {
        for (const kw of targetKeywords) {
          if (relU.includes(kw)) {
            score += 30;
            typeMatched = true;
            break;
          }
        }
      }

      // 如果路徑明確包含了「其他品類」的獨有關鍵字，且該關鍵字不屬於目前品類，則跳過此資料夾以防品類誤抓
      const hasConflictOther = otherProfileKeywords.some(okw => okw.length >= 3 && relU.includes(okw) && !targetKeywords.includes(okw));
      if (hasConflictOther && !typeMatched) {
        continue;
      }

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

      // Color matching with longest-match & prefix modifier protection
      let colorMatched = false;
      if (colorEn && relU.includes(colorEn.toUpperCase())) {
        score += 40;
        colorMatched = true;
      }

      const sortedColorAliases = [...colorAliases].sort((a, b) => b.length - a.length);
      for (const al of sortedColorAliases) {
        const alU = al.toUpperCase();
        const foundIdx = relU.indexOf(alU);
        if (foundIdx !== -1) {
          // 檢查前綴修飾字 (如 粉/酒/桃/紫/暗/淺/深)，若目標為純色則避免誤判衍生色
          const isPureZhColor = /^(紅|藍|綠|黃|棕|灰|黑|白|紫)$/.test(alU) || /^(紅色|藍色|綠色|黃色|棕色|灰色|黑色|白色|紫色)$/.test(alU);
          let isModified = false;
          if (isPureZhColor && foundIdx > 0) {
            const prevChar = relU[foundIdx - 1];
            if (/[粉酒桃紫玫暗淺深亮水湖]/.test(prevChar)) {
              isModified = true;
            }
          }

          if (!isModified) {
            score += 35;
            colorMatched = true;
            break;
          }
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

      if ((collMatched && colorMatched) || (typeMatched && (collMatched || colorMatched)) || score >= 50) {
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

  static isCategoryOrNonColor(folderName, templateProfiles = []) {
    if (!folderName) return true;
    const fUpper = folderName.toUpperCase().trim();
    const baseBadKeywords = [
      '圖片', 'PHOTO', 'IMAGE', 'IMAGES', 'BACK_LABEL', '背標', '尺寸', 'CHART', 'TABLE'
    ];
    for (const kw of baseBadKeywords) {
      if (fUpper.includes(kw)) return true;
    }

    const catKeywords = CoupangProcessor.getAllCategoryKeywords(templateProfiles);
    for (const kw of catKeywords) {
      const kwU = kw.toUpperCase().trim();
      if (fUpper === kwU || fUpper.includes(kwU)) return true;
    }

    if (/^[0-9._\s-]+$/.test(fUpper)) return true;
    return false;
  }

  static scanFolderStructure(photoFilesArray, templateProfiles = []) {
    const collections = new Set();
    const colors = new Set();
    const catKeywords = CoupangProcessor.getAllCategoryKeywords(templateProfiles);
    const escapedKws = catKeywords.map(k => k.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')).join('|');
    const catRegex = new RegExp(`(?:${escapedKws})\\s*(.+)$`, 'i');

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
      if (topDir && !this.isCategoryOrNonColor(topDir, templateProfiles)) {
        collections.add(topDir);
      }

      // 若只有一層資料夾且已被當作系列，則不作為顏色
      if (dirParts.length === 1 && collections.has(topDir)) {
        continue;
      }

      // 末端資料夾通常為顏色 (Color) 或單層商品名
      const leafDir = dirParts[dirParts.length - 1].trim();

      // 動態由品類關鍵字進行後綴截取 (如 "FIRENZE cat collar 天空藍" 或 "HERMITAGE RAINCOAT 黃色")
      let candColor = leafDir;
      const catMatch = candColor.match(catRegex);
      if (catMatch && catMatch[1].trim()) {
        candColor = catMatch[1].trim();
      }

      // 處理規格數字如 "Blue110_1.6_110_2.2"
      const dimMatch = candColor.match(/^([A-Za-z\u4e00-\u9fa5\s&（）\(\)]+?)(?:[0-9]+.*)?$/);
      if (dimMatch && dimMatch[1].trim()) {
        candColor = dimMatch[1].trim();
      }

      // 若過濾後不是品類詞、且不是已知的系列名稱，則加入顏色
      if (candColor && !this.isCategoryOrNonColor(candColor, templateProfiles) && !collections.has(candColor)) {
        colors.add(candColor);
      }
    }

    return {
      collections: Array.from(collections),
      colors: Array.from(colors)
    };
  }

  static extractMappingsFromExcel(workbook, headerRow = 3, rowStart = 4, filterColName = '中文背標', targetSheetName = null, templateProfiles = []) {
    let ws = null;
    if (targetSheetName) {
      const s = workbook.sheet(targetSheetName);
      if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
        ws = s;
      }
    }
    if (!ws) {
      const candidateNames = ['商品資料', '工作表1', 'Sheet1', 'Data', 'Sheet', '工作表'];
      for (const name of candidateNames) {
        const s = workbook.sheet(name);
        if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
          ws = s;
          break;
        }
      }
    }
    if (!ws) {
      const allSheets = workbook.sheets ? workbook.sheets() : [];
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
      if (bestSheet) ws = bestSheet;
    }
    if (!ws && targetSheetName) ws = workbook.sheet(targetSheetName);
    if (!ws) ws = workbook.sheet(0);
    if (!ws || !ws.usedRange()) return { collections: {}, colors: {} };

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

    const dummyProcessor = new CoupangProcessor([], null, {}, {}, templateProfiles);

    for (let r = rowStart; r <= totalRows; r++) {
      const zhName = getVal(r, colNameIdx);
      if (!zhName) continue;

      const rawColor = getVal(r, colColorIdx);
      const rawColl = getVal(r, colCollIdx);
      const rawSize = getVal(r, colSizeIdx);

      const parsed = dummyProcessor.parseChineseName(zhName, rawSize, { collection: rawColl, color: rawColor, size: rawSize });
      let zhColor = parsed.color || dummyProcessor.extractColorFromZhName(zhName, rawSize);
      const zhColl = (parsed.collection || '').replace(/系列$/g, '').trim();

      if (zhColor && CoupangProcessor.isCategoryOrNonColor(zhColor, templateProfiles)) {
        zhColor = '';
      }

      if (rawColl && zhColl && !CoupangProcessor.isCategoryOrNonColor(rawColl, templateProfiles)) {
        const cleanCollKey = rawColl.trim();
        if (!collectionMap[cleanCollKey]) collectionMap[cleanCollKey] = new Set();
        collectionMap[cleanCollKey].add(zhColl);
      }

      if (rawColor && zhColor && !CoupangProcessor.isCategoryOrNonColor(rawColor, templateProfiles)) {
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

  static mergeScannedAliases(existingCollections = {}, existingColors = {}, scannedFolder = { collections: [], colors: [] }, excelData = { collections: {}, colors: {} }, templateProfiles = []) {
    const newCollections = {};
    for (const [k, v] of Object.entries(existingCollections)) {
      if (!CoupangProcessor.isCategoryOrNonColor(k, templateProfiles)) {
        newCollections[k] = (v || []).filter(al => !CoupangProcessor.isCategoryOrNonColor(al, templateProfiles));
      }
    }

    const newColors = {};
    for (const [k, v] of Object.entries(existingColors)) {
      if (!CoupangProcessor.isCategoryOrNonColor(k, templateProfiles)) {
        newColors[k] = (v || []).filter(al => !CoupangProcessor.isCategoryOrNonColor(al, templateProfiles));
      }
    }

    const norm = s => (s || '').toUpperCase().replace(/[\（\(\）\)\s_-]/g, '').replace(/系列$/g, '');

    let newCollCount = 0;
    let newColorCount = 0;
    let matchedExcelCount = 0;

    // 1. 合併掃描到的系列 (Collections)
    for (const collName of scannedFolder.collections || []) {
      if (CoupangProcessor.isCategoryOrNonColor(collName, templateProfiles)) continue;
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
      if (CoupangProcessor.isCategoryOrNonColor(rawColl, templateProfiles)) continue;
      const nColl = norm(rawColl);
      if (!nColl) continue;

      const validZhColls = (zhColls || []).filter(zh => !CoupangProcessor.isCategoryOrNonColor(zh, templateProfiles));

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
      if (CoupangProcessor.isCategoryOrNonColor(colorName, templateProfiles)) continue;
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
      if (CoupangProcessor.isCategoryOrNonColor(rawColor, templateProfiles)) continue;
      const nCol = norm(rawColor);
      if (!nCol) continue;

      const validZhColors = (zhColors || []).filter(zh => !CoupangProcessor.isCategoryOrNonColor(zh, templateProfiles));

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

window.CoupangProcessor = CoupangProcessor;
