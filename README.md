# [coupang_excel_generator](https://github.com/enkaichang/coupang_excel_generator)

> 🚀 純前端、零依賴、跨平台的 Coupang (酷澎) 報價單與商品圖批量自動化生成系統。  
> 專為酷澎商品上架計畫設計，支援在瀏覽器端一鍵完成來源 Excel 欄位解析、品名規格拆解、圖片智慧配對、尺寸規範補正與報價單結構化輸出。

[![GitHub Repository](https://img.shields.io/badge/GitHub-coupang__excel__generator-181717?logo=github)](https://github.com/enkaichang/coupang_excel_generator)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version: v1.8.4](https://img.shields.io/badge/Version-v1.8.4-blue.svg)](https://github.com/enkaichang/coupang_excel_generator)

---

## 📑 目錄

- [系統特色](#-系統特色)
- [技術架構](#-技術架構)
- [專案結構](#-專案結構)
- [快速開始](#-快速開始)
- [功能與操作指引](#-功能與操作指引)
- [版本日誌 (Changelog)](#-版本日誌-changelog)
- [注意事項與常見問題](#-注意事項與常見問題)

---

## 🌟 系統特色

1. **純瀏覽器端運作（零安裝、高隱私）**：
   - 無需安裝 Python、Node.js 或任何後端環境，只要有現代瀏覽器（Chrome / Edge / Safari）即可直接執行。
   - 所有 Excel 解析與圖片縮放均在使用者本機記憶體完成，數據不經任何第三方伺服器，安全無外洩疑慮。
2. **完整樣式保留**：
   - 使用 xlsx-populate 讀寫 Excel，100% 保留標準報價單模板之所有公式、字型、背景色、框線、凍結窗格與欄寬。
3. **智慧屬性解析與配對**：
   - 自動自中文品名識別品牌、系列、顏色、尺寸與商品分類。
   - 多層級搜尋引擎：依據 SKU 代碼、系列英文名、顏色別名自動搜尋主圖、情境圖1、情境圖2及背標圖。
4. **一站式多模板設定檔與自動掃描精靈**：
   - 每個品類模板各自綁定獨立的匹配關鍵字、細分分類代碼、輸出子資料夾、Excel 範本檔與欄位對應（Fixed/Dynamic），一次設定完成。
   - 上傳新 Excel 模板時，系統自動掃描必填欄位並以精靈引導配置。
5. **圖片規格合規自動化**：
   - 自動檢測圖片寬度，主圖與情境圖自動等比放大並補白邊至 1000x1000，次要圖檔（尺寸表、背標）自動放大短邊至 1000px，確保符合電商上架規範。
6. **雙輸出模式**：
   - **一鍵存入資料夾 (推薦)**：採用 File System Access API，直接在使用者選定的資料夾內建立完整目錄結構並寫入檔案。
   - **ZIP 壓縮檔下載**：透過 JSZip 打包所有處理完畢之報價單與圖檔。
7. **可視化對照表設定與備份管理器**：
   - 視覺化管理多模板設定檔、系列中英對照表、顏色別名群組，並支援 IndexedDB 本地持久化與 JSON 備份檔匯出/匯入。

---

## 🛠 技術架構

- **核心語言**：HTML5, CSS3, Modern JavaScript (ES6+)
- **第三方核心函式庫**：
  - [xlsx-populate (v1.21.0)](https://github.com/dtjohnson/xlsx-populate)：瀏覽器端 Office Open XML 讀寫引擎。
  - [JSZip (v3.10.1)](https://stuk.github.io/jszip/)：瀏覽器端非同步 ZIP 打包與壓縮。
  - [FileSaver.js (v2.0.5)](https://github.com/eligrey/FileSaver.js/)：客戶端檔案下載介面。
- **瀏覽器原生 API**：
  - File System Access API (window.showDirectoryPicker)
  - IndexedDB API (突破 LocalStorage 容量限制持久儲存自訂模板)
  - HTML5 Canvas API (圖片縮放與格式轉換)
  - Web Storage API (localStorage)

---

## 📁 專案結構

```text
coupang_excel_generator/
├── index.html              # 主操作介面與 Modal 佈局
├── style.css               # 視覺設計、玻璃擬態樣式與響應式排版
├── app.js                  # 核心控制器、DOM 綁定、事件監聽與執行流程
├── README.md               # 專案說明與版本日誌
├── 欄位對應.md             # 來源與目標欄位對照、格式規範與檢查機制手冊
└── js/
    ├── storage_utils.js    # IndexedDB 儲存介面與設定檔備份匯出/匯入工具
    ├── default_config.js   # 預設模板設定檔、基準對照與系列/顏色字典
    ├── image_utils.js      # Canvas 圖片載入、尺寸檢測與等比縮放工具
    ├── my_family_processor.js # 核心業務邏輯（品名解析、分類判定、圖片搜尋與配對）
    └── templates.js        # 內建標準報價單 Excel 模板 Base64 編碼
`

---

## 🚀 快速開始

1. **開啟程式**：
   - 直接以瀏覽器開啟 index.html，或使用 VS Code 的 Live Server 擴充套件啟動。
2. **載入檔案**：
   - 將來源 Excel（如 .xlsx、.xls）拖曳至左側上傳區。
   - 點擊右側上傳區，選取本機之商品照片根目錄（如 Photo/）。
3. **設定處理範圍（可選）**：
   - 預設起始列為第 4 列。若需分批或測試局部資料，可輸入「結束列」或調整「起始列」。
4. **執行處理**：
   - 點擊「開始處理」，系統將即時更新進度條與診斷統計資訊（成功配對數、缺失圖片警告等）。
5. **匯出成果**：
   - 處理完成後，點擊「**一鍵存入資料夾**」（推薦）選取輸出目錄，或點擊「**打包下載 ZIP**」。

---

## ⚙️ 功能與操作指引

### 1. 對照表與多模板管理器
點擊右上角「**對照表設定**」按鈕，可開啟設定對話框：
- **模板與設定檔 (Templates & Profiles)**：一站式管理各品類模板之匹配關鍵字、細分分類代碼、輸出子資料夾、Excel 範本檔與獨立欄位對應（Fixed/Dynamic），支援上傳新模板並自動掃描必填欄位。
- **來源表設定 (Source)**：調整來源 Excel 工作表名稱、標題列、資料起始列與篩選欄位。
- **系列中英對照設定**：管理英文系列資料夾名稱與中文品名系列之對應關係，支援手動新增及一鍵自動掃描。
- **顏色別名設定**：管理顏色別名群組，支援手動維護與自動掃描擴充。
- **備份與還原**：提供「匯出備份 (JSON)」與「匯入設定 (JSON)」功能，輕鬆備份與跨裝置轉移完整設定。

---

## 📝 版本日誌 (Changelog)

### [v1.8.4] - 2026-08-14
#### 💄 全面純圖示化操作與文字精簡 (Pure Icon-Only Action Buttons & Clean UI)
- **全面移除圖示旁中文文字**：所有具圖示之操作按鈕（下載、設定、新增、重新掃描、重置、儲存、確認、取消等）全面精簡為純圖示按鈕（Pure Icon Button），徹底消除圖示與文字並存的雜亂感。
- **保留完整 Tooltip 懸浮提示**：每個圖示按鈕均保留詳細的 `title` 懸浮提示文字，兼具極簡乾淨的版面與清晰的使用指引。

---

### [v1.8.3] - 2026-08-14
#### 💄 介面 UX 與圖示化按鈕優化 (UI/UX & Icon Button Refinement)
- **按鈕視覺簡約與圖示化**：將刪除按鈕改為直觀的垃圾桶圖示、新增模板簡化為「+」圖示按鈕，並全面以 Material Icons（下載、儲存、新增等）增強按鈕語意與操作手感。
- **加強按鈕懸浮提示 (Tooltip)**：所有純圖示與簡化按鈕皆具備原生 `title` 懸浮提示，兼具精簡版面與無障礙操作體驗。

---

### [v1.8.2] - 2026-08-14
#### 💄 視覺圖示標準化與介面簡約優化 (Material Icons & Visual Simplification)
- **導入 Material Icons 視覺圖示庫**：全站所有 Emoji 符號全面改為 Google Material Icons，提升專業度與跨平台視覺一致性。
- **移除 suggest-tag 提示標籤**：移除補全視窗內多餘的提示標籤，推薦選項直接於下拉選單中清晰標註 `(推薦比對)`。

---

### [v1.8.1] - 2026-08-14
#### 🐛 同義詞精確推薦與品名欄位容錯優化 (Synonym Matching & Template Detection Refinement)
- **同義詞推薦二階段精確比對**：優先執行完全相等的關鍵字比對，確保欄位更名為「中文」時能精確推薦「商品名稱」，避免被「中文背標」等複合欄位搶先匹配。
- **模板偵測品名欄位擴充**：自動偵測所需模板時，擴大支援「中文」作為品名欄位，提升各類自訂表頭的相容性。

---

### [v1.8.0] - 2026-08-14
#### ✨ 來源 Excel 自動偵測模板與缺漏欄位智慧補全 (Auto Template Detection & Smart Column Mapping)
- **上傳自動偵測適用模板**：上傳來源 Excel 檔案後，系統即時掃描有效資料列，動態識別出此檔案實際需套用之模板設定檔（胸背帶、項圈、牽繩或自訂模板）。
- **缺漏欄位自動比對與智慧對應彈窗**：自動檢查來源 Excel 表頭是否具備命中的模板所需的所有動態對應欄位與篩選欄位。若有缺漏，自動跳出補全對話框，提供來源表頭下拉選單（含智慧推薦預選）、留空或設定固定值。
- **對應關係持久化**：支援勾選「記住此對應關係供日後使用」，確認後自動更新受影響模板之動態對應設定並持久化儲存。

---

### [v1.7.0] - 2026-08-13
#### 🚀 核心解析引擎泛用化與死碼清理 (Engine Generalization & Cleanup)
- **清理專案死碼與歷史寫死參數**：
  - 徹底移除 `targetCombos` 與 `zhColorMap` 等未使用的死碼與建構子引數。
  - 移除 `isCategoryOrNonColor` 中的 `NYLON`、`LEATHER`、`WITH CHAIN`、`ROPE LEASH` 等特定商品材質/黑名單詞彙，改由各 Profile 關鍵字與通用邏輯動態判定。
  - 移除針對 `X HARNESS` / `H HARNESS` 的硬編碼加分邏輯，以及工作表自動探測中的 `MYFAMILY` 等字樣。
- **擴充品名拆解彈性與來源欄位回退（Fallback）**：
  - 增強來源 Excel 表頭讀取，支援 `BRAND / 品牌` 欄位辨識。
  - 當來源品名結構非標準命名時，自動以 Excel 中的 `BRAND`、`COLLECTION`、`TYPE`、`COLOR`、`SIZE` 欄位值智慧回退補齊，確保各品類商品皆能正常解析與配對圖片。

---

### [v1.6.1] - 2026-08-13
#### 🐛 修復設定檔備份匯出錯誤 (Fix Config Export Reference Error)
- **修復匯出變數未定義**：修復匯出備份 JSON 時因未宣告 `sourceConfig`、`collectionAliases` 與 `colorAliases` 導致之執行例外。
- **即時同步介面設定**：匯出備份時自動同步並儲存當前介面各模板與對照表最新狀態。

---

### [v1.6.0] - 2026-08-13
#### 🚀 全新品類動態關鍵字驅動架構 (Profile-Driven Dynamic Category Engine)
- **全面動態化品類比對與過濾**：
  - 將品名拆解、照片目錄過濾與圖片評分引擎全面改由 Template Profiles 關鍵字動態驅動，徹底消除品類硬編碼。
  - 新增任意全新品類（如雨衣、名牌、衣服、玩具）時，系統自動支援品名斷詞、照片目錄過濾與圖片配對，無需修改程式碼。
  - 增強扁平單層目錄之屬性剝離與顏色前綴防護，徹底杜絕中文短詞（如紅色 vs 粉紅色）與複合顏色之匹配衝突。

---

### [v1.5.2] - 2026-08-13
#### ✨ 一站式模板與品類設定整合 (All-in-One Template & Category Management)
- **消除功能重複，統一整合至模板設定檔**：
  - 移除多餘的「品類與分類代碼設定」分頁，將品類名稱、匹配關鍵字、酷澎細分分類代碼、輸出子資料夾、Excel 範本與欄位映射完整整合至同一設定卡片。
  - 新增/編輯模板時可在同一個畫面一站式完成所有品類與欄位規則配置，無需在分頁間切換。

---

### [v1.5.1] - 2026-08-13
#### 💄 介面佈局優化 (UI Layout Optimization)
- **擴展對照表對話框寬度與排版空間**：
  - 將對照表設定視窗最大寬度由 720px 大幅擴展至 1080px，消除雙滾動條與欄位擠壓問題。
  - 優化多模板側邊欄與編輯面板之寬度比例（210px 側邊欄 + 自適應編輯區），並加寬固定/動態對應之目標欄位輸入框至 250px。
  - 改善頁籤（Tabs）與設定檔資訊列之間距與自適應換行體驗。

---

### [v1.5.0] - 2026-08-13
#### ✨ 多模板設定檔與自動掃描精靈升級 (Multi-Template Profiles & Setup Wizard)
- **多模板專屬設定檔架構 (Multi-Template Profiles)**：
  - 支援自訂多套品類設定檔（胸背帶、項圈、牽繩、服飾等），各自獨立綁定專屬的 Excel 模板檔案、酷澎細分分類代碼、動態欄位對應表 (Dynamic Mappings) 與固定填寫值 (Fixed Values)。
  - 每個設定檔具備獨立的匹配關鍵字與目標輸出子目錄，處理商品時自動分流至對應的報價單與輸出目錄。
- **自訂模板上傳與自動掃描精靈**：
  - 支援上傳自訂 .xlsx 模板檔案並直接儲存於瀏覽器 IndexedDB 中。
  - 上傳時自動偵測標題列並標註「必填欄位 (紅色)」與「選填欄位」，一鍵帶入智慧預設對應。
- **完整設定備份與還原 (JSON Import/Export)**：
  - 新增全域備份功能，可將所有設定檔、自訂模板二進位檔案、系列對照表與顏色別名完整匯出為單一 JSON 檔案，並支援跨裝置一鍵匯入還原。

---

### [v1.4.7] - 2026-08-13
#### ⚡️ 圖片尺寸調整與處理流程優化 (Image Resizing and Processing Optimization)
- **主圖與情境圖 1000x1000 補白邊 (Canvas Resize & Pad)**：
  - 商品主圖與情境圖一律等比縮放並在 1000x1000 白色畫布上置中（使用 #FFFFFF 補齊白邊），完全符合酷澎商品圖片規範。
- **尺寸圖與背標圖短邊解析度保證**：
  - 尺寸規格表與背標圖維持短邊至少 1000px 等比放大，確保文字與條碼細節清晰可辨。

---

### [v1.4.6] - 2026-08-12
#### ⚡️ 進階記憶體與檔案處理效能優化 (Performance & Memory Optimization)
- **零殘留記憶體管理**：
  - 優化大型照片資料夾與多工作表之記憶體佔用，改進 DOM 渲染與事件解綁機制。

---

### [v1.4.5] - 2026-08-12
#### 🔧 跨格式檔名比對與正則強化 (Filename Normalization & Regular Expression Hardening)
- **智慧檔名防呆機制**：
  - 強化商品品名與圖檔規格之正則拆解，支援更多非標準空格、全形標點與特殊符號之容錯。

---

### [v1.4.4] - 2026-08-12
#### 🐛 報價單欄位格式與維度修復 (Coupang Template Compatibility Fix)
- **OOXML 結構補全**：
  - 補齊 xlsx-populate 輸出時遺失之 <dimension> 標籤，解決酷澎上架系統解析異常問題。

---

### [v1.0.0] - 2026-08-11
- 初版發布：支援基礎 Excel 解析、多品類報價單產生與圖檔整理。