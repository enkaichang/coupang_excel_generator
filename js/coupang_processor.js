class CoupangProcessor {
  constructor(photoFilesArray, excelFileBuffer, templateProfiles = []) {
    // 支援彈性引數 (相容舊呼叫)
    if (Array.isArray(arguments[2]) && arguments[2].length > 0 && arguments[2][0]?.template_file_name) {
      templateProfiles = arguments[2];
    } else if (Array.isArray(arguments[4])) {
      templateProfiles = arguments[4];
    }
    this.photoFiles = photoFilesArray || [];
    this.excelBuffer = excelFileBuffer;
    this.templateProfiles = templateProfiles || [];
    this.categoryRules = this.templateProfiles;

    this.folderMap = new Map();
    for (const file of (this.photoFiles || [])) {
      const normPath = (file.webkitRelativePath || file.name || '').replace(/\\/g, '/');
      if (/(?:^|\/)(result|output|__macosx|\.git)(?:\/|$)/i.test(normPath)) {
        continue;
      }
      const pathParts = normPath.split('/');
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

  static getDefaultProfiles() {
    if (typeof window !== 'undefined' && window.AppConfig && typeof window.AppConfig.getDefaultProfiles === 'function') {
      return window.AppConfig.getDefaultProfiles();
    }
    if (typeof require !== 'undefined') {
      try {
        const defaultCfg = require('./default_config.js');
        if (defaultCfg && typeof defaultCfg.getDefaultProfiles === 'function') {
          return defaultCfg.getDefaultProfiles();
        }
      } catch (e) {}
    }
    return [];
  }

  static getAllCategoryKeywords(templateProfiles = []) {
    const profiles = (templateProfiles && templateProfiles.length > 0)
      ? templateProfiles
      : CoupangProcessor.getDefaultProfiles();

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

  static getUtils() {
    if (typeof window !== 'undefined' && window.SharedUtils) return window.SharedUtils;
    if (typeof require !== 'undefined') {
      try { return require('./shared_utils.js'); } catch (e) {}
    }
    return null;
  }

  normStr(s) {
    if (!s) return '';
    const utils = CoupangProcessor.getUtils();
    if (utils) return utils.cleanStrForMatching(s);
    return String(s).normalize('NFKC').toUpperCase().replace(/[\u3000\u00A0\s\r\n\(\)（）\-_*\/\\#\[\]]/g, '').trim();
  }

  isKeywordMatch(text, kw) {
    if (!text || !kw) return false;
    const utils = (typeof window !== 'undefined' && window.SharedUtils) ? window.SharedUtils : (typeof require !== 'undefined' ? require('./shared_utils.js') : null);
    if (utils) {
      return utils.matchToken(text, kw, { boundaryMatch: true, strictNoPrefix: true });
    }
    const textUpper = String(text).normalize('NFKC').trim().toUpperCase();
    const kwUpper = String(kw).normalize('NFKC').trim().toUpperCase();
    return textUpper.includes(kwUpper);
  }

  calculateMainImageScore(fileName) {
    if (!fileName) return 0;
    const fn = fileName;
    const fnL = fn.toLowerCase();

    // 嚴格僅抓「主圖」、「main」、「主圖1」
    const isExplicitMain = fn.includes('主圖') || fnL.includes('main') || fn.includes('主圖1');
    if (isExplicitMain) {
      return 100;
    }
    return 0;
  }

  calculateSceneImageScore(fileName, sceneIndex = 1) {
    if (!fileName) return 0;
    const fn = fileName;
    const fnL = fn.toLowerCase();
    const nameNoExt = fn.replace(/\.[^/.]+$/, '').trim().toLowerCase();

    const targetNum = String(sceneIndex);
    // 嚴格僅抓「情境圖 {n}」、「情境圖{n}」、「情境{n}」、「scene{n}」、「{n}」
    const isExplicitScene =
      fn.includes(`情境圖 ${targetNum}`) ||
      fn.includes(`情境圖${targetNum}`) ||
      fn.includes(`情境${targetNum}`) ||
      fnL.includes(`scene${targetNum}`) ||
      fnL.includes(`scene_${targetNum}`) ||
      fnL.includes(`scene ${targetNum}`) ||
      nameNoExt === targetNum ||
      (nameNoExt === '情境圖' && sceneIndex === 1) ||
      (nameNoExt === '情境' && sceneIndex === 1);

    if (isExplicitScene) {
      return 100;
    }
    return 0;
  }

  getTargetTemplateAndCategory(prodName = '', typeStr = '') {
    const typeUpper = (typeStr || '').toUpperCase().trim();
    const nameUpper = (prodName || '').toUpperCase().trim();
    const fullText = (typeUpper + ' ' + nameUpper).trim();

    const profiles = (this.templateProfiles && this.templateProfiles.length > 0)
      ? this.templateProfiles
      : CoupangProcessor.getDefaultProfiles();

    let bestProfile = null;
    let highestScore = 0;

    for (const p of profiles) {
      const keywords = (p.keywords || [p.name]).map(k => (k || '').toString().trim().toUpperCase()).filter(Boolean);
      let profileScore = 0;

      for (const k of keywords) {
        let kwScore = 0;

        // 1. Type 欄位精確完全相符 (最高優先權)
        if (typeUpper && typeUpper === k) {
          kwScore = Math.max(kwScore, 1000 + k.length * 10);
        }
        // 2. Type 欄位單詞吻合 (例如 "COOLING JACKET" 中包含單詞 "JACKET")
        else if (typeUpper && this.isKeywordMatch(typeUpper, k)) {
          kwScore = Math.max(kwScore, 500 + k.length * 10);
        }

        // 3. 中文品名完全相符
        if (nameUpper && nameUpper === k) {
          kwScore = Math.max(kwScore, 800 + k.length * 10);
        }
        // 4. 中文品名/全文包含 (依字串長度加權，長詞優先避免短詞攔截)
        else if (this.isKeywordMatch(nameUpper, k) || this.isKeywordMatch(fullText, k)) {
          kwScore = Math.max(kwScore, 100 + k.length * 10);
        }

        if (kwScore > profileScore) {
          profileScore = kwScore;
        }
      }

      if (profileScore > highestScore) {
        highestScore = profileScore;
        bestProfile = p;
      }
    }

    if (bestProfile && highestScore > 0) {
      const tmplId = bestProfile.id || bestProfile.template_id || bestProfile.template_type || 'LEASH';
      const tmplType = (bestProfile.template_type || bestProfile.id || 'LEASH').toUpperCase();
      const tmplFile = bestProfile.template_file_name || bestProfile.template_name || (tmplType === 'HARNESS' ? '商品報價單_胸背帶.xlsx' : '商品報價單_項圈 牽繩.xlsx');
      const outputFile = tmplFile.replace(/\.xlsx$/i, '_auto_generate.xlsx');
      const utils = CoupangProcessor.getUtils();
      const targetSubfolder = (utils && typeof utils.getTemplateSubfolder === 'function')
        ? utils.getTemplateSubfolder(bestProfile, this.templateProfiles)
        : (bestProfile.name || (tmplType === 'HARNESS' ? '胸背帶' : '項圈 牽繩'));
      return {
        template_id: tmplId,
        template_type: tmplType,
        template_file: tmplFile,
        output_file: outputFile,
        category: bestProfile.category_name || (tmplType === 'HARNESS' ? '寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)' : '寵物用品>犬貓通用>項圈/伸縮牽繩>項圈 (66025)'),
        target_subfolder: targetSubfolder,
        matched_rule: bestProfile.name || '',
        profile: bestProfile,
        unmatched: false
      };
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
    return "";
  }

  extractChineseKeywords(prodName, brand = '', color = '', size = '') {
    if (!prodName) return [];
    const utils = CoupangProcessor.getUtils();
    let s = (utils ? utils.normalizeKey(prodName) : String(prodName).replace(/[\r\n\t]+/g, ' ').normalize('NFKC').trim());
    if (brand) {
      const escapedBrand = brand.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      s = s.replace(new RegExp(escapedBrand, 'gi'), ' ');
    }
    s = s.replace(/[\（\(\）\)]/g, ' ');
    if (color) {
      const escapedColor = String(color).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      s = s.replace(new RegExp(escapedColor, 'gi'), ' ');
    }
    if (size) {
      const normSize = String(size).normalize('NFKC').trim();
      const escapedSize = size.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      const escapedNormSize = normSize.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      s = s.replace(new RegExp(`(?:^|[\\s_\\-\\(\\[])(?:${escapedSize}|${escapedNormSize})(?:$|[\\s_\\-\\)\\]])`, 'gi'), ' ');
    }

    const rawTokens = s.split(/[\s_+\/\-]+/).map(t => t.trim()).filter(t => t.length >= 2);
    const keywords = new Set(rawTokens);

    // 通用子詞分解：長度 >= 4 的詞段自動萃取 2~3 字子詞，完全消除廠商專屬硬編碼
    for (const token of rawTokens) {
      if (token.length >= 4) {
        for (let len = 2; len <= 3; len++) {
          for (let i = 0; i <= token.length - len; i++) {
            keywords.add(token.substr(i, len));
          }
        }
      }
    }
    return Array.from(keywords);
  }

  calculateBackLabelScore(fileName, normTargetSize, zhKeywords = [], colorZh = '', collZh = '', collEn = '', prodZhName = '') {
    if (!fileName) return 0;
    const utils = (typeof window !== 'undefined' && window.SharedUtils) ? window.SharedUtils : (typeof require !== 'undefined' ? require('./shared_utils.js') : null);
    const fn = fileName;
    const dotIdx = fn.lastIndexOf('.');
    const nameRaw = dotIdx !== -1 ? fn.substring(0, dotIdx) : fn;
    const nameNoExt = nameRaw.normalize('NFKC').toUpperCase();
    const nameNormalized = nameRaw.replace(/_/g, ' ');
    const nameNoUnderscore = nameRaw.replace(/_/g, '');
    const nameClean = this.normStr(nameRaw);
    let sizeScore = 0;

    if (normTargetSize === 'UNISIZE' || normTargetSize === '單一尺寸' || normTargetSize === 'ONE' || normTargetSize === 'ONE SIZE') {
      if (nameNoExt.includes('單一尺寸') || nameNoExt.includes('UNISIZE') || nameNoExt.includes('_ONE') || nameNoExt.includes(' ONE') || nameNoExt === '背標_ONE' || nameNoExt === '背標_單一尺寸') {
        sizeScore = 100;
      } else if (nameNoExt.includes('背標') || nameNoExt.includes('LABEL')) {
        sizeScore = 80;
      }
    } else if (normTargetSize) {
      const escaped = normTargetSize.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      // 1. 支援標準專屬命名直接命中: 背標_{尺寸} / 背標-{尺寸}
      const regexDirectBackLabel = new RegExp(`^背標[_\-\\s]?${escaped}$`, 'i');

      if (regexDirectBackLabel.test(nameNoExt)) {
        sizeScore = 110;
      } else if (utils ? utils.matchToken(nameNoExt, normTargetSize, { boundaryMatch: true }) : new RegExp(`(?:^|[\\s_\\-\\(\\[])${escaped}(?:$|[\\s_\\-\\)\\]]|_?背標|\\.[a-zA-Z0-9]+$)`, 'i').test(nameNoExt)) {
        sizeScore = 100;
      } else {
        const cleanTargetDim = normTargetSize.replace(/CM|MM/g, '').replace(/[\s\/\-_*xX\.]+/g, '_');
        const cleanNameDim = nameNoExt.replace(/CM|MM/g, '').replace(/[\s\/\-_*xX\.]+/g, '_');
        if (cleanTargetDim.length >= 3 && cleanNameDim.includes(cleanTargetDim)) {
          sizeScore = 90;
        }
      }
    } else {
      // 單一尺寸或無尺寸欄位商品（normTargetSize 為空）：尋找專屬背標且未標註其他尺寸代碼 (1個明確 Fallback)
      if (nameNoExt.includes('背標') || nameNoExt.includes('LABEL')) {
        const hasOtherSize = /(?:^|[\s_\-\(\[])([0-9]*X*[SML]+|[0-9]+-[0-9.]+|[0-9]+\/[0-9]+mm)(?:$|[\s_\-\)\]])/i.test(nameNoExt);
        if (!hasOtherSize) {
          sizeScore = 90;
        }
      }
    }

    if (sizeScore === 0) return 0;

    let totalScore = sizeScore;

    // 圖片經統一正規化後比對是否包含完整中文品名 (支援包含完整中文品名之圖檔)
    if (prodZhName) {
      const cleanProdName = this.normStr(prodZhName);
      if (cleanProdName.length >= 3 && nameClean.includes(cleanProdName)) {
        totalScore += 50;
      }
    }

    // 中文品名核心詞交集加分 (每個吻合關鍵詞 +25 分，支援底線轉換、去底線與直接匹配)
    for (const kw of zhKeywords) {
      if (kw.length >= 2 && (nameRaw.includes(kw) || nameNormalized.includes(kw) || nameNoUnderscore.includes(kw) || nameClean.includes(kw))) {
        totalScore += 25;
      }
    }

    // 中文顏色命中加分 (+30 分，前綴安全防護)
    if (colorZh && (utils ? (utils.matchToken(nameRaw, colorZh, { strictNoPrefix: true }) || utils.matchToken(nameNoUnderscore, colorZh, { strictNoPrefix: true })) : nameRaw.includes(colorZh))) {
      totalScore += 30;
    }

    // 系列命中加分 (+20 分 / +15 分)
    if (collZh && (utils ? utils.matchToken(nameRaw, collZh) : nameRaw.includes(collZh))) {
      totalScore += 20;
    }
    if (collEn && (utils ? utils.matchToken(nameNoExt, collEn) : nameNoExt.includes(collEn.toUpperCase()))) {
      totalScore += 15;
    }

    return totalScore;
  }

  parseChineseName(prodName, excelSize = "", rawHints = {}) {
    const brand = (rawHints && rawHints.brand) ? String(rawHints.brand).trim() : '';
    const collection = (rawHints && rawHints.collection) ? String(rawHints.collection).trim() : '';
    const type = (rawHints && rawHints.type) ? String(rawHints.type).trim() : '';
    const color = (rawHints && (rawHints.zhColor || rawHints.color)) ? String(rawHints.zhColor || rawHints.color).trim() : '';
    const size = (rawHints && rawHints.size) ? String(rawHints.size).trim() : (excelSize ? String(excelSize).trim() : '');

    return {
      success: true,
      brand,
      collection: (collection || '').replace(/系列$/g, '').trim(),
      type,
      color,
      size,
      rawName: prodName,
      lastNameToken: size
    };
  }

  getLocalImagesForProduct(parsedData, rawHints = {}) {
    const utils = CoupangProcessor.getUtils();
    const collEn = (rawHints.collection || '').trim();
    const collZh = collEn ? '' : (parsedData.collection || '').replace(/系列$/g, '').trim();
    const brandName = (rawHints.brand || parsedData.brand || '').trim();
    const collectionCandidates = (collEn ? [collEn] : (collZh ? [collZh] : []))
      .filter(c => !brandName || !c.toUpperCase().includes(brandName.toUpperCase()));

    const typZh = (parsedData.type || '').trim();
    const typEn = (rawHints.type || '').trim();
    const typeStr = (typZh + ' ' + typEn).toUpperCase();

    // 顏色：優先使用來源 Excel 之「中文顏色」，其次使用中文品名解析出之顏色，並以英文 Color 作為備用
    const colorZh = (rawHints.zhColor || rawHints.colorZh || parsedData.color || '').trim();
    const colorEn = (rawHints.color || rawHints.colorEn || '').trim();
    const labelTargetStr = (parsedData.size || rawHints.size || parsedData.lastNameToken || '').trim();
    const sizeStr = labelTargetStr.toUpperCase();
    const skuStr = (rawHints.sku || '').toString().trim();

    // 萃取品名核心中文關鍵字（排除品牌、顏色、尺寸干擾，並拆解複合名詞）
    const prodZhName = String(parsedData.rawName || rawHints.name || (parsedData.brand + ' ' + (parsedData.collection || '') + ' ' + (parsedData.type || ''))).replace(/[\r\n\t]+/g, ' ').trim();
    const zhKeywords = this.extractChineseKeywords(prodZhName, parsedData.brand || rawHints.brand || '', colorZh, labelTargetStr);
    if (collZh && !zhKeywords.includes(collZh)) zhKeywords.push(collZh);

    // 動態取得所屬品類 Profile 之所有關鍵字與排除衝突之其他品類關鍵字
    const targetProfileInfo = this.getTargetTemplateAndCategory(prodZhName, typeStr);
    const profiles = (this.templateProfiles && this.templateProfiles.length > 0)
      ? this.templateProfiles
      : CoupangProcessor.getDefaultProfiles();

    let currentProfile = profiles.find(p => p.id === targetProfileInfo.template_id);
    if (!currentProfile) {
      currentProfile = profiles.find(p => p.id === targetProfileInfo.template_type) || profiles[0];
    }

    const targetKeywords = (currentProfile?.keywords || [currentProfile?.name || ''])
      .map(k => (k || '').toString().toUpperCase().trim())
      .filter(Boolean);

    // 2. 直通 Excel 的 TYPE：提取細分品類 Token (動態排除品牌、系列、大品類關鍵字，100% 跨廠商通用無硬編碼)
    const excludeTokens = new Set();
    if (brandName) {
      brandName.toUpperCase().split(/[\s_+\/\-]+/).forEach(t => { if (t.length >= 2) excludeTokens.add(t); });
    }
    for (const c of collectionCandidates) {
      if (c) c.toUpperCase().split(/[\s_+\/\-]+/).forEach(t => { if (t.length >= 2) excludeTokens.add(t); });
    }
    for (const kw of targetKeywords) {
      if (kw) kw.toUpperCase().split(/[\s_+\/\-]+/).forEach(t => { if (t.length >= 2) excludeTokens.add(t); });
    }

    const typeTokensEn = typEn.split(/[\s_+\/\-]+/)
      .map(t => t.toUpperCase().trim())
      .filter(t => t.length >= 2 && !excludeTokens.has(t));

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
    const normTargetSize = sizeStr ? sizeStr.normalize('NFKC').replace(/號/g, '').trim().toUpperCase() : '';

    for (const [dirPath, files] of this.folderMap.entries()) {
      const relU = dirPath.toUpperCase();
      let score = 0;

      // 0. 目錄路徑完整包含中文品名加分 (+60 分，長詞/精準品名目錄最高優先權)
      const dirClean = this.normStr(dirPath);
      const cleanProd = this.normStr(prodZhName);
      if (cleanProd.length >= 3 && dirClean.includes(cleanProd)) {
        score += 60;
      }

      // 1. Type Profile matching: 動態判定路徑是否符合該品類關鍵字
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

      // 2. Collection matching (直通 Excel COLLECTION 欄位比對)
      let collMatched = false;
      const utils = (typeof window !== 'undefined' && window.SharedUtils) ? window.SharedUtils : (typeof require !== 'undefined' ? require('./shared_utils.js') : null);

      for (const cal of collectionCandidates) {
        if (utils ? utils.matchToken(relU, cal, { boundaryMatch: true }) : relU.includes(cal.toUpperCase())) {
          score += 60;
          collMatched = true;
          break;
        } else if (relU.includes(cal.toUpperCase())) {
          score += 40;
          collMatched = true;
          break;
        }
      }
      if (!collMatched && collectionCandidates.length > 0) {
        for (const cal of collectionCandidates) {
          if (files.some(f => utils ? utils.matchToken(f.name, cal, { boundaryMatch: true }) : f.name.toUpperCase().includes(cal.toUpperCase()))) {
            score += 25;
            collMatched = true;
            break;
          }
        }
      }

      // 3. Product Sub-Type matching (直通 Excel TYPE 欄位細分比對，如 APRON vs VEST, TOWEL vs POCKET)
      for (const t of typeTokensEn) {
        if (utils ? utils.matchToken(relU, t, { boundaryMatch: true }) : relU.includes(t)) {
          score += 45;
        }
      }

      // 4. Color matching: 優先比對中文顏色 (Photo/Collection/Type/中文顏色/)
      let colorMatched = false;
      if (colorZh) {
        if (utils ? utils.matchToken(relU, colorZh, { boundaryMatch: true, strictNoPrefix: true }) : relU.includes(colorZh.toUpperCase())) {
          score += 40;
          colorMatched = true;
        } else if (files.some(f => utils ? utils.matchToken(f.name, colorZh, { boundaryMatch: true, strictNoPrefix: true }) : f.name.includes(colorZh))) {
          score += 30;
          colorMatched = true;
        }
      }

      // 備用：比對英文顏色
      if (!colorMatched && colorEn && (utils ? utils.matchToken(relU, colorEn, { boundaryMatch: true }) : relU.includes(colorEn.toUpperCase()))) {
        score += 35;
        colorMatched = true;
      }

      // 5. 尺寸精準邊界加分
      if (normTargetSize) {
        if (files.some(f => utils ? utils.matchToken(f.name, normTargetSize, { boundaryMatch: true }) : f.name.includes(normTargetSize))) {
          score += 15;
        }
      }

      // 6. 中文品名特徵關鍵字加分 (支援底線圖檔與複合分詞)
      if (zhKeywords.length > 0) {
        const hitKwCount = zhKeywords.filter(kw => kw.length >= 2 && files.some(f => f.name.includes(kw) || f.name.replace(/_/g, ' ').includes(kw))).length;
        if (hitKwCount > 0) {
          score += (hitKwCount * 15);
        }
      }

      // 7. 背標精確存在性與品名語意綜合加分（解除 120 封頂，真實回饋圖檔精確度）
      let maxLblScore = 0;
      for (const f of files) {
        const fl = f.name.toLowerCase();
        if (f.name.includes('主圖') || fl.startsWith('main') || f.name.includes('情境') || f.name.includes('尺寸') || fl.includes('chart')) continue;
        const lblScore = this.calculateBackLabelScore(f.name, normTargetSize, zhKeywords, colorZh, collZh, collEn, prodZhName);
        if (lblScore > maxLblScore) maxLblScore = lblScore;
      }

      // 解除 120 分上限限制，將背標精準匹配分數如實計入目錄總分
      score += maxLblScore;

      // 嚴格判定：當商品有明確指定系列 (Collection) 時，必須命中系列 (collMatched) 或分數達到高信賴度門檻才可視為有效候選資料夾，杜絕跨系列誤配
      const collValid = (collectionCandidates.length === 0) || collMatched;
      if (collValid && (collMatched || score >= 120)) {
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

    if (!matchedFolder && matchedFolder !== '') {
      return { main: null, sc1: null, sc2: null, chart: null, label: null };
    }

    // 扁平目錄防呆檢查 (若使用者將所有圖片上傳至根目錄，無子資料夾可區分)
    const isFlatDir = (this.folderMap.size === 1 && this.folderMap.has('')) || matchedFolder === '';

    let mainImg = null;
    let sc1 = null;
    let sc2 = null;
    let chart = null;
    let label = null;

    let bestMainScore = 0;
    let bestSc1Score = 0;
    let bestSc2Score = 0;

    const isImgFile = (name) => /\.(jpe?g|png|webp|jfif|avif)$/i.test(name || '');

    for (const f of matchedFiles) {
      const fn = f.name;
      const fnL = fn.toLowerCase();
      if (!isImgFile(fn)) continue;

      // 扁平目錄防呆：若無子目錄區分，必須符合商品名稱或條碼特徵
      if (isFlatDir) {
        const matchesProduct = (skuStr && fn.includes(skuStr)) ||
          (colorZh && (utils ? utils.matchToken(fn, colorZh, { strictNoPrefix: true }) : fn.includes(colorZh))) ||
          (collZh && (utils ? utils.matchToken(fn, collZh) : fn.includes(collZh))) ||
          (zhKeywords.length > 0 && zhKeywords.some(k => fn.includes(k)));
        if (!matchesProduct) continue;
      }

      // 尺寸圖挑選
      if (!chart && (fn.includes('尺寸圖') || fn.includes('尺寸表') || fnL.includes('chart') || fn.includes('尺寸規格'))) {
        chart = f;
      }

      // 主圖挑選 (嚴格只抓主圖/main/主圖1)
      const mScore = this.calculateMainImageScore(fn);
      if (mScore > bestMainScore && mScore >= 100) {
        bestMainScore = mScore;
        mainImg = f;
      }

      // 情境圖 1 挑選 (嚴格只抓情境圖 1/情境圖1/情境1/scene1/1)
      const s1Score = this.calculateSceneImageScore(fn, 1);
      if (s1Score > bestSc1Score && s1Score >= 100) {
        bestSc1Score = s1Score;
        sc1 = f;
      }

      // 情境圖 2 挑選 (嚴格只抓情境圖 2/情境圖2/情境2/scene2/2)
      const s2Score = this.calculateSceneImageScore(fn, 2);
      if (s2Score > bestSc2Score && s2Score >= 100) {
        bestSc2Score = s2Score;
        sc2 = f;
      }
    }

    // 1. 背標精確匹配（結合中文品名核心詞交集評分）
    const backLabelCandidates = matchedFiles.filter(f => {
      const fn = f.name;
      const fl = fn.toLowerCase();
      if (!isImgFile(fn)) return false;
      if (fn.includes('主圖') || fl.startsWith('main') || fn.includes('主圖1')) return false;
      if (fn.includes('情境') || fl.includes('scene')) return false;
      if (fn.includes('尺寸圖') || fn.includes('尺寸表') || fl.includes('chart') || fn.includes('尺寸規格')) return false;
      if (fn === '.DS_Store' || fl.endsWith('.db')) return false;
      return true;
    });

    let bestScore = 0;
    for (const f of backLabelCandidates) {
      const fScore = this.calculateBackLabelScore(f.name, normTargetSize, zhKeywords, colorZh, collZh, collEn, prodZhName);
      if (fScore > bestScore && fScore >= 90) {
        bestScore = fScore;
        label = f;
      }
    }

    // 2. 尺寸圖單一明確回退 (Fallback 1)：若當前顏色資料夾無尺寸圖，向上層/同系列資料夾尋找
    if (!chart && matchedFolder) {
      const lastSlashIdx = matchedFolder.lastIndexOf('/');
      const parentPath = lastSlashIdx !== -1 ? matchedFolder.substring(0, lastSlashIdx) : '';
      for (const [dPath, files] of this.folderMap.entries()) {
        const isParentMatch = parentPath && parentPath.length > 0 && dPath.startsWith(parentPath);
        const isCollMatch = (collZh && (utils ? utils.matchToken(dPath, collZh) : dPath.toUpperCase().includes(collZh.toUpperCase()))) ||
                            (collEn && (utils ? utils.matchToken(dPath, collEn) : dPath.toUpperCase().includes(collEn.toUpperCase())));
        if (isParentMatch || isCollMatch) {
          for (const f of files) {
            const fn = f.name;
            const fnL = fn.toLowerCase();
            if (!isImgFile(fn)) continue;
            if (fn.includes('尺寸圖') || fn.includes('尺寸表') || fnL.includes('chart') || fn.includes('尺寸規格')) {
              chart = f;
              break;
            }
          }
        }
        if (chart) break;
      }
    }

    // 恪守「寧可空白，也不要抓錯」：若未找到對應圖檔即維持 null，絕不任意冒充
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
      const relPath = (file.webkitRelativePath || file.name || '').replace(/\\/g, '/');
      if (!relPath || /(?:^|\/)(result|output|__macosx|\.git)(?:\/|$)/i.test(relPath)) continue;

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

  static extractMappingsFromExcel(workbook, headerRow = 3, rowStart = 4, filterColName = '', targetSheetName = null, templateProfiles = []) {
    const utils = CoupangProcessor.getUtils();
    const ws = utils
      ? utils.getSourceSheet(workbook, targetSheetName, headerRow)
      : (workbook.sheet(targetSheetName) || workbook.sheet(0));
    if (!ws || !ws.usedRange()) return { collections: {}, colors: {} };

    const usedRange = ws.usedRange();
    if (!usedRange) return { collections: {}, colors: {} };
    const maxCol = usedRange.endCell().columnNumber();
    const totalRows = usedRange.endCell().rowNumber();

    const headers = {};
    for (let c = 1; c <= maxCol; c++) {
      const val = ws.cell(headerRow, c).value();
      if (val) {
        const raw = val.toString().trim();
        const norm = utils ? utils.normalizeKey(raw) : raw;
        if (!(raw in headers)) headers[raw] = c;
        if (!(norm in headers)) headers[norm] = c;
      }
    }

    // 尋找目標欄位 index (支援獨立英文與中文欄位比對)
    const colColorEnIdx = headers['COLOR'] || headers['Color'] || headers['Colour'];
    const colColorZhIdx = headers['中文顏色'] || headers['顏色'] || headers['顏色(中)'];
    const colCollEnIdx = headers['COLLECTION'] || headers['Collection'];
    const colCollZhIdx = headers['系列'] || headers['中文系列'] || headers['系列名稱'];

    const getVal = (r, colIdx) => {
      if (!colIdx) return '';
      const cell = ws.cell(r, colIdx);
      if (utils) return String(utils.getCellValue(cell)).trim();
      const v = cell.value();
      if (v === null || v === undefined) return '';
      if (typeof v === 'object' && typeof v.text === 'function') return v.text().trim();
      return v.toString().trim();
    };

    const collectionMap = {}; // { "HERMITAGE": Set(["隱士", ...]) }
    const colorMap = {};      // { "CAMEL": Set(["可可棕", ...]) }

    for (let r = rowStart; r <= totalRows; r++) {
      const rawCollEn = getVal(r, colCollEnIdx);
      const rawCollZh = getVal(r, colCollZhIdx);
      const rawColorEn = getVal(r, colColorEnIdx);
      const rawColorZh = getVal(r, colColorZhIdx);

      const zhColl = (rawCollZh || '').replace(/系列$/g, '').trim();
      const zhColor = (rawColorZh || '').trim();

      if (rawCollEn && zhColl && rawCollEn !== zhColl && !CoupangProcessor.isCategoryOrNonColor(rawCollEn, templateProfiles) && !CoupangProcessor.isCategoryOrNonColor(zhColl, templateProfiles)) {
        const cleanCollKey = rawCollEn.trim();
        if (!collectionMap[cleanCollKey]) collectionMap[cleanCollKey] = new Set();
        collectionMap[cleanCollKey].add(zhColl);
      }

      if (rawColorEn && zhColor && rawColorEn !== zhColor && !CoupangProcessor.isCategoryOrNonColor(rawColorEn, templateProfiles) && !CoupangProcessor.isCategoryOrNonColor(zhColor, templateProfiles)) {
        const cleanColorKey = rawColorEn.trim();
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

if (typeof module !== 'undefined' && module.exports) {
  module.exports = CoupangProcessor;
}
if (typeof window !== 'undefined') {
  window.CoupangProcessor = CoupangProcessor;
}
