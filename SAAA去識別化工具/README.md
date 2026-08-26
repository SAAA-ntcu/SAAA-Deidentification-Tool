# SAAA 學生學習能力檢測去識別化工具 (SAAA De-identification Tool)

[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live%20Demo-brightgreen?style=flat&logo=github)](https://liangchengyu-ntcu.github.io/SAAA-Deidentification-Tool/)
[![Bootstrap 5](https://img.shields.io/badge/Bootstrap-5.3.3-purple?logo=bootstrap)](https://getbootstrap.com/)
[![SweetAlert2](https://img.shields.io/badge/SweetAlert2-Popup-red)](https://sweetalert2.github.io/)
[![100% Client-Side](https://img.shields.io/badge/Security-100%25%20Offline%20%26%20Client--Side-blue?logo=shield)](https://github.com/)

> 🔒 **100% 本機端瀏覽器離線運算** · **免安裝** · **零個資上傳** · **完美支援新竹縣與各縣市學力檢測多結構 Excel**

---

## 📖 專案簡介

本工具專為各級學校、教育局端及學術研究人員打造，提供直覺的網頁圖形介面，用於快速處理大量學生評量成績、作答反應、班平均與試題分析檔案之**去識別化（脫敏）**工作。

### 🌟 核心特色
1. **🛡️ 100% 本機離線安全**：採用純前端 HTML5 + Web Worker/ArrayBuffer 引擎，所有檔案轉換均在您電腦的瀏覽器記憶體中運算，絕不上傳任何伺服器，拔掉網路線亦可高速執行。
2. **🚫 Windows 0 阻擋**：無任何 `.exe` 執行檔，徹底杜絕 Windows SmartScreen 藍色警告與各家防毒軟體誤判。
3. **📂 智慧辨識五大多型態評量報表**：
   - **個人成績**（每檔含數十個獨立學生 Sheet，支援工作表標籤與內部卡片自動脫敏）
   - **作答反應**（學生每題原始選答反應清單，支援跨科學生代碼完美對齊）
   - **答對率與各校等級資料**（個人 PR/向度答對率與多層次表頭學校彙整表）
   - **班平均**（全縣各校各班級平均表現表）
   - **試題分析**（多工作表題項難易度、鑑別度與選項分析）
4. **🔬 三大去識別化模式**：
   - **研究分析模式**：學生姓名轉為跨科一致假名代碼（如 `STU_112_0001`），自動產出 `mapping_table.xlsx` 對照總表。
   - **公開資料模式**：徹底剔除姓名與座號欄位，不可逆且零個資留存（適用於政府資料開放、新聞稿發布）。
   - **校內遮罩模式**：中文姓名智慧遮罩（如 `林*偉`），Sheet 標籤遮罩為合法名稱（`林*偉成績`），保留原表結構。
5. **🗜️ 支援 .zip 一鍵解壓與目錄鏡像打包**：直接拖入整個資料夾或 Zip 包，自動處理並打包下載。
6. **✨ 現代化視覺 UI**：採用 Bootstrap 5 響應式元件與 SweetAlert2 質感彈跳提醒。

---

## 🚀 快速開始 (3 步驟)

1. 用任何現代瀏覽器（Google Chrome、Microsoft Edge、Safari 等）開啟 [**`index.html`**](index.html)。
2. 將評量資料夾或 `.zip` 壓縮檔**直接拖入網頁中間的虛線拖曳區**。
3. 選擇所需的去識別化模式與學校別名，點擊 **「🚀 開始執行去識別化並打包下載」**！

---

## 📁 檔案結構

```
📁 SAAA去識別化工具/
│
├── 📄 index.html                  # 標準 HTML5 語意化首頁 (雙擊即開)
├── 📄 README.md                   # 專案說明與操作手冊
│
├── 📁 css/                        # 樣式目錄 (本地免連網)
│   ├── 📄 bootstrap.min.css       # Bootstrap 5 核心樣式
│   ├── 📄 sweetalert2.min.css     # SweetAlert2 彈跳視窗樣式
│   └── 📄 style.css               # 自訂 UI 樣式
│
├── 📁 js/                         # 業務邏輯模組
│   └── 📄 app.js                  # 去識別化引擎、Mapper 與 UI 控制器
│
└── 📁 lib/                        # 開源核心函式庫 (本地免連網)
    ├── 📄 bootstrap.bundle.min.js # Bootstrap 互動函式庫
    ├── 📄 sweetalert2.all.min.js  # SweetAlert2 彈跳視窗引擎
    ├── 📄 xlsx.full.min.js        # SheetJS Excel 讀寫引擎
    └── 📄 jszip.min.js            # JSZip 壓縮打包引擎
```

---

## 🌐 部署至 GitHub Pages

本工具為純靜態網站（Static Web App），可直接推送到 GitHub 並開啟 **GitHub Pages** 免費線上運作：

1. 將本專案推送到 GitHub Repository。
2. 進入 Repository 的 **Settings** $\to$ **Pages**。
3. 在 **Build and deployment** 下方的 **Source** 選擇 **Deploy from a branch** 或透過 **GitHub Actions** 自動化部署。
4. 部署完成後即可獲得專屬線上使用網址，隨開隨用！

---

## 🔒 資安與隱私合規聲明

- 本專案程式碼 **100% 於 Client-Side (瀏覽器端) 執行**。
- 無後端 API、無資料庫、無任何第三方追蹤代碼（No Google Analytics / Tracking Cookies）。
- 使用者處理之評量資料與學生個人隱私資訊均不離開使用者的電腦硬體。
