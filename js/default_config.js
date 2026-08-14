window.AppConfig = {
  getDefaultSourceConfig: function() {
    return {
      file_path: "商品資料.xlsx",
      sheet_name: "Sheet1",
      header_row: 3,
      row_start: 4,
      filter_column: "中文背標"
    };
  },

  getBaselineMappings: function() {
    return {
      dynamic: {
        "商品名稱": "中文品名",
        "商品條碼": "EAN",
        "酷澎進價 (含稅)": "酷澎進價",
        "建議酷澎售價": "台灣訂價",
        "每單位包裝尺寸(mm)": "每單位包裝尺寸\n(mm)",
        "每單位包裝重量(g)": "每單位包裝重量\n(g)"
      },
      fixed: {
        "數量": "1個",
        "包裝上方\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "包裝下方\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "包裝左側\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "包裝右側\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "是否應稅": "應稅(5%)",
        "製造廠商": "",
        "供貨方式": "官方代理商",
        "是否為進口商品": "進口商品",
        "單箱商品入數": "1",
        "商品保存天數": "1095",
        "是否為易碎品": "不適用",
        "國內製造商或負責商名稱": "緯豪實業有限公司",
        "法定種類": "",
        "原產地/國": "中國",
        "注意事項/備註欄": "請參考產品實際包裝所示"
      },
      systemFields: [
        "細分商品種類",
        "商品名稱",
        "品牌",
        "系列",
        "顏色",
        "尺寸",
        "商品正面(主要圖片）",
        "商品側拍或情境圖 1\n(消費者可見圖片）",
        "商品側拍或情境圖 2\n(消費者可見圖片）",
        "商品側拍或情境圖 3\n(消費者可見圖片）",
        "商品側拍或情境圖 4\n(消費者可見圖片）",
        "商品詳細說明圖集\n(消費者可見圖片）",
        "商品實際中文背標圖\n(內部上架審核用)",
        "商品條碼照\n(內部上架審核用)"
      ]
    };
  },

  getDefaultProfiles: function() {
    const baseline = this.getBaselineMappings();
    return [
      {
        id: 'HARNESS',
        name: '胸背帶',
        keywords: ['HARNESS', '胸背帶', '背帶'],
        template_type: 'HARNESS',
        template_file_name: '商品報價單_胸背帶.xlsx',
        category_name: '寵物用品>狗用品>牽繩/胸背帶>胸背帶 (66030)',
        subfolder: '胸背帶',
        is_builtin: true,
        field_mappings: {
          dynamic: { ...baseline.dynamic },
          fixed: { ...baseline.fixed }
        }
      },
      {
        id: 'COLLAR',
        name: '項圈',
        keywords: ['COLLAR', '項圈'],
        template_type: 'LEASH',
        template_file_name: '商品報價單_項圈 牽繩.xlsx',
        category_name: '寵物用品>犬貓通用>項圈/伸縮牽繩>項圈 (66025)',
        subfolder: '項圈',
        is_builtin: true,
        field_mappings: {
          dynamic: { ...baseline.dynamic },
          fixed: { ...baseline.fixed }
        }
      },
      {
        id: 'LEASH',
        name: '牽繩',
        keywords: ['LEASH', '牽繩', '拉繩'],
        template_type: 'LEASH',
        template_file_name: '商品報價單_項圈 牽繩.xlsx',
        category_name: '寵物用品>犬貓通用>項圈/伸縮牽繩>牽繩 (66027)',
        subfolder: '牽繩',
        is_builtin: true,
        field_mappings: {
          dynamic: { ...baseline.dynamic },
          fixed: { ...baseline.fixed }
        }
      }
    ];
  },

  getDefaultConfig: function() {
    const baseline = this.getBaselineMappings();
    return {
      source: this.getDefaultSourceConfig(),
      field_mappings: {
        dynamic: { ...baseline.dynamic },
        fixed: { ...baseline.fixed }
      }
    };
  },

  getDefaultCollectionAliases: function() {
    return {
      'HERMITAGE': ['HERMITAGE', '隱士'],
      'FIRENZE': ['FIRENZE', '佛羅倫斯'],
      'AMALFI': ['AMALFI', '阿瑪菲'],
      'ASCOT': ['ASCOT', '雅士閣', '雅士谷', '阿斯科特'],
      'GINEVRA': ['GINEVRA', '吉內瓦', '日內瓦'],
      'MemopetID': ['MemopetID', 'Memopet'],
      'Milano': ['Milano', '米蘭'],
      'Monza': ['Monza', '蒙札', '蒙扎'],
      'ROYAL': ['ROYAL', '皇家'],
      'SAINT TROPEZ': ['SAINT TROPEZ', '聖特羅佩'],
      'Tucson': ['Tucson', '土桑']
    };
  },

  getDefaultColorAliases: function() {
    return {
      'TARTAN LILAC & CREAM': ['TARTAN LILAC & CREAM', 'LILAC', '丁香紫格紋', '丁香紫', '紫格紋'],
      'BLACK & OCHRE': ['BLACK & OCHRE', 'BLACK & OCHRE（Tan）', '經典黑/土黃色', '經典黑', '土黃色', '經典黑/大地黃'],
      'BROWN & TARTAN': ['BROWN & TARTAN', '可可棕格紋', '棕格紋', '可可棕/格紋色', '英倫格紋'],
      'BROWN & TURQUOISE': ['BROWN & TURQUOISE', 'BROWN & TURQUOISE（light blue）', 'BROWN & LIGHTBLUE', '可可棕/土耳其藍', '可可棕/天空藍'],
      'MANDARIN': ['MANDARIN', 'TANGERINE', 'TENGERINE', '活力橘', '橘色', '橘', '暖陽橘', '蜜柑橘'],
      'BLACK': ['BLACK', '經典黑', '黑色', '黑'],
      'BROWN': ['BROWN', 'CAMEL', '可可棕', '焦糖棕', '棕色', '棕'],
      'PINK': ['PINK', 'PINK（Wisteria）', '櫻花粉', '粉紅色', '粉色', '粉', '柔粉', '甜蜜粉', '蜜桃粉'],
      'RED': ['RED', 'BORDEAUX', '勃根地紅', '酒紅', '紅色', '紅', '赤紅', '經典紅', '波爾多紅'],
      'FUCHSIA': ['FUCHSIA', '桃紅', '紫桃紅'],
      'GREEN': ['GREEN', 'APPLE GREEN', 'LIME', '蘋果綠', '綠色', '綠', '萊姆綠', '森林綠', '螢光綠'],
      'BLUE': ['BLUE', 'LIGHT BLUE', 'NAVY', 'DARK BLUE', '天空藍', '藍色', '藍', '靜謐藍', '海軍藍', '深海藍', '經典藍', '蔚藍'],
      'PURPLE': ['PURPLE', 'AMETHYST', 'LIGHT PURPLE', '粉紫', '粉紫色', '紫色', '紫', '葡萄紫', '丁香紫'],
      'GREY': ['GREY', 'GRAY', 'GREY & BLACK', '摩登灰', '莫蘭迪灰', '灰色', '灰', '雲朵灰']
    };
  },

  get: function() {
    const saved = localStorage.getItem('coupang_config') || localStorage.getItem('my_family_config');
    const def = this.getDefaultConfig();
    if (!saved) return def;
    try {
      const parsed = JSON.parse(saved);
      const dyn = { ...def.field_mappings.dynamic, ...(parsed.field_mappings?.dynamic || {}) };
      if (dyn['顏色'] === 'COLOR') delete dyn['顏色'];
      if (dyn['尺寸'] === 'SIZE') delete dyn['尺寸'];
      return {
        source: { ...def.source, ...(parsed.source || {}) },
        field_mappings: {
          fixed: { ...def.field_mappings.fixed, ...(parsed.field_mappings?.fixed || {}) },
          dynamic: dyn
        }
      };
    } catch (e) {
      return def;
    }
  },

  getDefaultCategoryRules: function() {
    return [];
  },

  getCategoryRules: function() {
    return [];
  },

  getCollectionAliases: function() {
    const saved = localStorage.getItem('coupang_collection_aliases') || localStorage.getItem('my_family_collection_aliases');
    return saved ? JSON.parse(saved) : this.getDefaultCollectionAliases();
  },

  getColorAliases: function() {
    const saved = localStorage.getItem('coupang_color_aliases') || localStorage.getItem('my_family_color_aliases');
    return saved ? JSON.parse(saved) : this.getDefaultColorAliases();
  }
};
