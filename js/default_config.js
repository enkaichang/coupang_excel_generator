const AppConfig = {
  getDefaultSourceConfig: function() {
    return {
      file_path: "商品資料.xlsx",
      sheet_name: "Sheet1",
      header_row: 3,
      row_start: 4,
      filter_column: "",
      brand_fixed: "",
      manufacturer_fixed: "",
      collection_column: "COLLECTION",
      type_column: "TYPE",
      color_column: "中文顏色",
      size_column: "SIZE",
      fixed: {
        "包裝上方\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "包裝下方\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "包裝左側\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "包裝右側\n(內部上架審核用)": "嚴正聲明-本商品無外包裝照片",
        "是否應稅": "應稅(5%)",
        "供貨方式": "官方代理商",
        "是否為進口商品": "進口商品",
        "法定種類": "TW_General",
        "國內製造商或負責商名稱": "緯豪實業有限公司"
      }
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
        "單箱商品入數": "1",
        "商品保存天數": "1095",
        "是否為易碎品": "不適用",
        "原產地/國": "中國",
        "注意事項/備註欄": "請參考產品實際包裝所示"
      },
      systemFields: [
        "細分商品種類",
        "商品名稱",
        "品牌",
        "製造廠商",
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
        is_builtin: true,
        field_mappings: {
          dynamic: { ...baseline.dynamic },
          fixed: { ...baseline.fixed }
        }
      },
      {
        id: 'TOY',
        name: '玩具及訓練工具',
        keywords: ['TOY', 'DUMMY', 'FLOATING TOY', 'CALM', '玩具', '玩偶', '安撫玩偶', '訓練球', '訓練玩具', '拉扯玩具'],
        template_type: 'ACCESSORIES',
        template_file_name: '商品報價單_生活與訓練配件.xlsx',
        category_name: '寵物用品>犬用玩具>咬咬/拉扯玩具 (66040)',
        is_builtin: true,
        field_mappings: {
          dynamic: { ...baseline.dynamic },
          fixed: { ...baseline.fixed }
        }
      },
      {
        id: 'BALL',
        name: '玩具球',
        keywords: ['BALL', '彈力球', '球'],
        template_type: 'ACCESSORIES',
        template_file_name: '商品報價單_生活與訓練配件.xlsx',
        category_name: '寵物用品>犬用玩具>球類玩具 (66041)',
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
    return {};
  },

  getDefaultColorAliases: function() {
    return {};
  },

  get: function() {
    const saved = localStorage.getItem('coupang_config');
    const def = this.getDefaultConfig();
    if (!saved) return def;
    try {
      const parsed = JSON.parse(saved);
      const dyn = { ...def.field_mappings.dynamic, ...(parsed.field_mappings?.dynamic || {}) };
      if (dyn['顏色'] === 'COLOR') delete dyn['顏色'];
      if (dyn['尺寸'] === 'SIZE') delete dyn['尺寸'];
      return {
        source: {
          ...def.source,
          ...(parsed.source || {}),
          fixed: { ...def.source.fixed, ...(parsed.source?.fixed || {}) }
        },
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
    return {};
  },

  getColorAliases: function() {
    return {};
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = AppConfig;
}
if (typeof window !== 'undefined') {
  window.AppConfig = AppConfig;
}
