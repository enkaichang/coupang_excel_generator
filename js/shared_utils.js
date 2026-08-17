/**
 * Coupang Excel Generator - Shared Utilities
 * 統一字串正規化、全半形符號相容、數值防呆 (防 NaN)、檔名安全過濾與共用工作表探測工具
 */
(function() {
  const SharedUtils = {
    /**
     * 標準化表頭或比對鍵值：
     * 1. 統一轉為 NFKC (相容全形英數)
     * 2. 移除 BOM (\uFEFF)、零寬字元 (\u200B)
     * 3. 將全形空白 (\u3000)、不換行空白 (\u00A0) 轉為標準空白
     * 4. 統一換行符號為 \n
     * 5. 將全形括號 （） 轉為半形 ()
     */
    normalizeKey: function(str) {
      if (str === null || str === undefined) return '';
      return String(str)
        .normalize('NFKC')
        .replace(/[\uFEFF\u200B]/g, '')
        .replace(/[\u3000\u00A0]/g, ' ')
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/（/g, '(')
        .replace(/）/g, ')')
        .trim();
    },

    /**
     * 深度標準化字串 (用於模糊比對與別名查找)：
     * 移除所有空白、括號、換行、底線、破折號、特殊標記並轉大寫
     */
    cleanStrForMatching: function(str) {
      if (str === null || str === undefined) return '';
      return String(str)
        .normalize('NFKC')
        .replace(/[\uFEFF\u200B]/g, '')
        .replace(/[\u3000\u00A0\s\r\n\(\)（）\-_*\/\\#\[\]]/g, '')
        .toUpperCase()
        .trim();
    },

    /**
     * 安全檔案名稱清理：
     * 1. 移除 Windows / Unix 禁用字元: \ / : * ? " < > | 以及控制字元 \r \n \t \0
     * 2. 移除 Windows 檔名末尾禁止的句點 . 與空白
     * 3. 確保不回傳空字串
     */
    sanitizeFilename: function(name, fallbackName = '未命名檔案') {
      if (!name) return fallbackName;
      let s = String(name)
        .normalize('NFKC')
        .replace(/[\\/:*?"<>|\r\n\t\0]/g, '_')
        .replace(/[.\s]+$/, '')
        .trim();
      return s || fallbackName;
    },

    /**
     * 條碼 / EAN / SKU 格式化：
     * 徹底防止科學記號 (8.05E+12) 與小數點 .0，確保以純數字文字輸出
     */
    formatBarcode: function(val) {
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
    },

    /**
     * 價格 (進價 / 建議售價) 數值清理：
     * 過濾 NT$、$、元、千分位逗號，轉換為純數值 (Number)
     * 若無法解析為有效正數則回傳空字串 "" (落實「寧願空著也不要出錯」)
     */
    cleanPrice: function(val) {
      if (val === null || val === undefined || val === '') return '';
      if (typeof val === 'number') {
        return (!isNaN(val) && val >= 0) ? val : '';
      }
      const s = String(val).normalize('NFKC').replace(/[\u3000\u00A0\s,NT$元]/gi, '').trim();
      if (!s) return '';
      const num = parseFloat(s);
      return (!isNaN(num) && num >= 0) ? num : '';
    },

    /**
     * 包裝尺寸清理與換算：
     * 1. 支援乘號相容: *, x, X, ×, ＊, 乘
     * 2. 自動偵測 cm 單位並乘以 10 換算為整數 mm
     * 3. 輸出標準 長*寬*高 (如 120*80*25)
     * 4. 若格式無效或無法解析為 3 個正整數，直接回傳原字串或空字串，絕不輸出 NaN
     */
    cleanPackagingDimension: function(val) {
      if (val === null || val === undefined || val === '') return '';
      let s = String(val).normalize('NFKC').trim();
      if (!s) return '';

      const isCm = /cm|公分/i.test(s);
      let cleanS = s.replace(/(mm|cm|公厘|公分)/gi, '').trim();
      cleanS = cleanS.replace(/[\u00D7\u2715\u2716\uFF0AxX＊*乘\s]+/g, '*');
      
      const parts = cleanS.split('*').map(p => p.trim()).filter(Boolean);
      if (parts.length === 3) {
        const nums = parts.map(p => {
          const n = parseFloat(p);
          if (isNaN(n) || n < 0) return null;
          return isCm ? Math.round(n * 10) : Math.round(n);
        });

        if (nums.every(n => n !== null && !isNaN(n))) {
          return nums.join('*');
        }
      }

      // 若非 3 項目標準格式，進行基礎清理，避免殘留危險字元
      return s.replace(/[\r\n\t]/g, ' ').trim();
    },

    /**
     * 包裝重量清理：
     * 1. 過濾 g, kg, 公克, 約 等字樣
     * 2. kg 自動乘以 1000 換算為整數公克 (g)
     * 3. 若無法轉換為有效正整數，回傳空字串 "" (徹底杜絕 NaN 寫入 Excel)
     */
    cleanPackagingWeight: function(val) {
      if (val === null || val === undefined || val === '') return '';
      if (typeof val === 'number') {
        return (!isNaN(val) && val >= 0) ? Math.round(val) : '';
      }
      let s = String(val).normalize('NFKC').trim();
      if (!s) return '';

      const isKg = /kg|公斤/i.test(s);
      const numMatch = s.replace(/,/g, '').match(/[0-9]+(?:\.[0-9]+)?/);
      if (!numMatch) return '';

      const num = parseFloat(numMatch[0]);
      if (isNaN(num) || num < 0) return '';

      return isKg ? Math.round(num * 1000) : Math.round(num);
    },

    /**
     * 儲存格安全讀值：
     * 容錯 RichText、公式物件與 Error 物件，避免非預期型別崩潰
     */
    getCellValue: function(cell) {
      if (!cell) return '';
      const val = cell.value();
      if (val === null || val === undefined) return '';
      if (typeof val === 'object') {
        if (typeof val.text === 'function') return val.text();
        if (val.text) return String(val.text);
        if (val.result !== undefined) return String(val.result);
        if (val.error) return '';
      }
      return val;
    },

    /**
     * 通用邊界與字詞精確比對器 (Anti-containment & Exact Token Matcher)
     * 支援：
     * 1. 英文/數字 ASCII 單詞邊界檢查 (防止 CAT 誤中 CATEGORY、S 誤中 XS、RED 誤中 DARKRED)
     * 2. 中文符號/中英交界邊界檢查 (支援 _, -, /, 空白, 括號, 檔名副檔名邊界, BATHROBE黑, COLLAR甜杏)
     * 3. 中文修飾前綴防護 (strictNoPrefix: 防止 紅色 誤中 粉紅色、暗紅色、酒紅色)
     * 
     * @param {string} text 待搜尋的全文 (如檔名或路徑)
     * @param {string} target 要尋找的關鍵字/Token (如顏色、尺寸、系列、品名詞)
     * @param {object} options 設定選項
     *   - exactOnly: boolean (是否要求全文完全相符)
     *   - boundaryMatch: boolean (預設 true, 是否要求處於單詞/符號邊界)
     *   - strictNoPrefix: boolean (預設 true, 是否排除中文修飾前綴詞)
     *   - caseSensitive: boolean (預設 false)
     * @returns {boolean}
     */
    matchToken: function(text, target, options = {}) {
      if (!text || !target) return false;
      const exactOnly = options.exactOnly || false;
      const boundaryMatch = options.boundaryMatch !== undefined ? options.boundaryMatch : true;
      const strictNoPrefix = options.strictNoPrefix !== undefined ? options.strictNoPrefix : true;
      const caseSensitive = options.caseSensitive || false;

      const normText = this.normalizeKey(text);
      const normTarget = this.normalizeKey(target);
      if (!normText || !normTarget) return false;

      const t = caseSensitive ? normText : normText.toUpperCase();
      const k = caseSensitive ? normTarget : normTarget.toUpperCase();

      if (exactOnly) {
        return t === k;
      }

      if (t === k) return true;
      if (k.length > t.length) return false;

      const isTargetAscii = /^[\x00-\x7F]+$/.test(k);
      const escapedK = k.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');

      if (boundaryMatch) {
        if (isTargetAscii) {
          const asciiBoundaryRegex = new RegExp(`(?:^|[^A-Z0-9])${escapedK}(?:$|[^A-Z0-9]|_?背標|\\.[a-zA-Z0-9]+$)`, caseSensitive ? '' : 'i');
          if (asciiBoundaryRegex.test(t)) {
            return true;
          }
        } else {
          // 中文邊界：前後為符號、空白、英數字交界或字串起訖點
          const zhBoundaryRegex = new RegExp(`(?:^|[^a-zA-Z0-9\u4e00-\u9fa5]|[a-zA-Z0-9])${escapedK}(?:$|[^a-zA-Z0-9\u4e00-\u9fa5]|[a-zA-Z0-9]|_?背標|\\.[a-zA-Z0-9]+$)`, caseSensitive ? '' : 'i');
          if (zhBoundaryRegex.test(t)) {
            // 需再確認前綴字元非中文修飾前綴
            const matchIdx = t.indexOf(k);
            if (matchIdx !== -1) {
              const prevChar = matchIdx > 0 ? t[matchIdx - 1] : '';
              const prefixChars = '粉暗酒深淺亮桃水寶草青嫩土鐵米墨灰紫橘橙軍鮮雙金銀微超抗防透涼';
              if (!strictNoPrefix || !prefixChars.includes(prevChar)) {
                return true;
              }
            }
          }
        }
      }

      // 上下文前綴/後綴檢驗
      let startIdx = 0;
      while ((startIdx = t.indexOf(k, startIdx)) !== -1) {
        let isValidMatch = true;

        if (startIdx > 0) {
          const prevChar = t[startIdx - 1];
          if (isTargetAscii && /[A-Z0-9]/i.test(prevChar)) {
            isValidMatch = false;
          }
          if (strictNoPrefix) {
            const prefixChars = '粉暗酒深淺亮桃水寶草青嫩土鐵米墨灰紫橘橙軍鮮雙金銀微超抗防透涼';
            if (prefixChars.includes(prevChar)) {
              isValidMatch = false;
            }
          }
        }

        if ((startIdx + k.length) < t.length) {
          const nextChar = t[startIdx + k.length];
          if (isTargetAscii && /[A-Z0-9]/i.test(nextChar)) {
            isValidMatch = false;
          }
        }

        if (isValidMatch) {
          return true;
        }

        startIdx += k.length;
      }

      return false;
    },

    isExactTokenMatch: function(text, token) {
      return this.matchToken(text, token, { boundaryMatch: true, strictNoPrefix: true });
    },

    /**
     * 來源工作表自動探測 (共用統一函式)
     */
    getSourceSheet: function(workbook, targetSheetName = null, headerRow = 3) {
      if (!workbook) return null;

      if (targetSheetName) {
        const s = workbook.sheet(targetSheetName);
        if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
          return s;
        }
      }

      const candidateNames = ['商品資料', '工作表1', 'Sheet1', 'Data', 'Sheet', '工作表'];
      for (const name of candidateNames) {
        const s = workbook.sheet(name);
        if (s && s.usedRange() && s.usedRange().endCell().rowNumber() > headerRow) {
          return s;
        }
      }

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
      if (bestSheet) return bestSheet;

      if (targetSheetName) {
        const s = workbook.sheet(targetSheetName);
        if (s) return s;
      }
      for (const name of candidateNames) {
        const s = workbook.sheet(name);
        if (s) return s;
      }
      return workbook.sheet(0);
    }
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = SharedUtils;
  }
  if (typeof window !== 'undefined') {
    window.SharedUtils = SharedUtils;
  }
})();
