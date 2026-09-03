# SAAA 學生學習能力檢測去識別化工具 (SAAA De-identification Tool)

[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live%20Demo-brightgreen?style=flat&logo=github)](https://saaa-ntcu.github.io/SAAA-Deidentification-Tool/)
[![Build & Release Desktop Apps](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/actions/workflows/release-desktop.yml/badge.svg)](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/actions/workflows/release-desktop.yml)
[![CI Check](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/actions/workflows/ci.yml/badge.svg)](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/actions/workflows/ci.yml)
[![Tauri v2](https://img.shields.io/badge/Tauri-v2.0-blue?logo=tauri)](https://v2.tauri.app/)
[![Bootstrap 5](https://img.shields.io/badge/Bootstrap-5.3.3-purple?logo=bootstrap)](https://getbootstrap.com/)
[![SweetAlert2](https://img.shields.io/badge/SweetAlert2-Popup-red)](https://sweetalert2.github.io/)
[![100% Client-Side](https://img.shields.io/badge/Security-100%25%20Offline%20%26%20Client--Side-blue?logo=shield)](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool)

> 🔒 **100% 本機離線運算** · **零個資上傳** · **支援跨平台原生桌面應用程式 (Windows .msi/.exe, macOS .dmg) 與純瀏覽器免安裝運行** · **專門適配 SAAA 學力檢測 7 大類評量 Excel**

---

## 💻 全平台使用方式 (Windows / macOS / 網頁免安裝)

本工具為各級學校教師、資訊組長與研究人員提供三種使用管道，所有管道皆為 **100% 本地運算、絕不上傳任何個資至雲端**：

### 🪟 Windows 教師使用指南 (推薦桌面版)

Windows 老師可至 [GitHub Releases 發布頁面](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/releases) 下載專屬安裝檔：

| 檔案類型 | 推薦對象 | 特點說明 |
|---|---|---|
| **`*.exe` (NSIS 安裝檔)** | 國小 / 國中一般教師 | 繁體中文精靈引導，雙擊「下一步」即可安裝完成，自動建立桌面捷徑與開始功能表圖示。 |
| **`*.msi` (Windows Installer)** | 學校資訊組長 / 系統管理員 | 具備標準 Windows 封裝規格，支援 Active Directory / 群組原則 (GPO) 批次派送至電腦教室。 |

> 💡 **不想安裝任何軟體的 Windows 老師？**  
> 請直接使用「方案三：免安裝純網頁離線版」或「方案四：GitHub Pages 線上版」，用 Chrome 或 Edge 即可直接開啟使用！

---

### 🍎 macOS 教師使用指南 (原生桌面版)

1. 至 [Releases 頁面](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/releases) 下載 `SAAA_Deidentification_Tool_macOS.dmg`。
2. 雙擊開啟 `.dmg`，將 **「SAAA去識別化工具」** 拖曳至 **「Applications (應用程式)」** 資料夾。
3. 於應用程式清單中點擊啟動即可。原生 Apple Silicon / Intel 雙架構支援，僅約 2.3 MB 超輕量。

---

### 🌐 方案三：免安裝純網頁離線版 (ZIP 下載，拔掉網路線可用)

1. 至 [Releases 頁面](https://github.com/SAAA-ntcu/SAAA-Deidentification-Tool/releases) 下載 `SAAA_Deidentification_Tool_Offline.zip`。
2. 解壓縮後，**直接點兩下打開 `index.html`**。
3. 所有相依套件（Bootstrap, SweetAlert2, SheetJS, JSZip）皆已本機化於資料夾中，在完全無網路的環境下亦可流暢運行。

---

### 🚀 方案四：GitHub Pages 線上直接使用 (零下載、零安裝)

直接點擊前往：**[SAAA 去識別化工具 線上版 (GitHub Pages)](https://saaa-ntcu.github.io/SAAA-Deidentification-Tool/)**  
直接將評量資料夾或 `.zip` 拖入網頁，即可在瀏覽器記憶體中完成去識別化並儲存下載。

---

## 📁 儲存庫結構導覽 (Repository Structure)

```
📁 SAAA-Deidentification-Tool/
│
├── 📁 tauri-app/                    # 🖥️ Tauri v2 跨平台原生桌面應用程式
│   ├── 📁 src-tauri/                # Rust 後端核心
│   │   ├── 📄 Cargo.toml            # Rust 專案相依性定義
│   │   ├── 📄 tauri.conf.json       # Tauri 視窗、圖示、Windows NSIS 繁中配置
│   │   ├── 📁 icons/                # 高解析度跨平台 App 圖示 (.icns, .ico, .png)
│   │   └── 📁 src/
│   │       ├── 📄 main.rs           # 桌面程式進入點
│   │       └── 📄 lib.rs            # save_bytes 原生系統另存新檔對話框 IPC 命令
│   ├── 📁 frontend/                 # 桌面版嵌入之前端靜態資源
│   ├── 📄 package.json              # 前端建置與 Tauri CLI 腳本
│   └── 📄 package-lock.json
│
├── 📁 SAAA去識別化工具/             # 🌐 核心網頁工具（GitHub Pages 自動部署根目錄）
│   ├── 📄 index.html                # 語意化首頁 (Bootstrap 5 + SweetAlert2)
│   ├── 📄 README.md                 # 評量資料型態規格與去識別化技術手冊
│   ├── 📁 css/                      # 離線 CSS 樣式庫
│   ├── 📁 js/                       # 核心脫敏運算引擎 (app.js)
│   └── 📁 lib/                      # 嵌入式引擎 (xlsx, jszip, bootstrap, sweetalert2)
│
├── 📁 .github/workflows/            # ⚙️ GitHub Actions 自動化 CI/CD 流程
│   ├── 📄 release-desktop.yml       # 🪟🍎 自動編譯 Windows (.msi/.exe) 與 macOS (.dmg) 並發布 Release
│   ├── 📄 deploy.yml                # 🚀 自動部署 SAAA 網頁工具至 GitHub Pages
│   └── 📄 ci.yml                    # 🛡️ JavaScript 語法檢測與敏感資料防護
│
├── 📄 .gitignore                    # 🛡️ 嚴格隱私過濾規則 (自動排除所有學生個資與建置暫存)
└── 📄 README.md                     # 專案總覽文件
```

---

## ⚙️ GitHub Actions 自動編譯與發布流程 (CI/CD)

本儲存庫配置了全自動化的雲端編譯流程：

1. **自動發布 Release**：
   - 每當在 GitHub 推送版本標籤（例如 `git tag v2.0.0 && git push origin v2.0.0`），將自動啟動：
     - **Windows Runner (`windows-latest`)**：自動安裝 Rust 與 WiX/NSIS，編譯產出 Windows `.msi` 與 `.exe`。
     - **macOS Runner (`macos-latest`)**：自動編譯產出原生 `.app` 並製作安裝映像檔 `.dmg`。
     - **Ubuntu Runner (`ubuntu-latest`)**：自動打包純網頁離線壓縮包 `SAAA_Deidentification_Tool_Offline.zip`。
     - **發布作業 (`publish-release`)**：自動彙整所有平台檔案，一鍵建立 GitHub Release 並附帶各平台下載安裝指南！
2. **手動測試觸發 (`workflow_dispatch`)**：
   - 即使未打標籤，開發者亦可隨時至 Actions 頁面手動點擊「Run workflow」，系統建置完成後會將產物儲存於 Artifacts 提供下載測試。

---

## 📊 去識別化核心功能與處理機制

本工具精確解析 SAAA 固定格式的 **7 大類 Excel 報表**：
1. **個人成績**（各班卡片式診斷表，精確保留等級描述，清除座號與卡片標籤姓名）
2. **作答反應**（解析學生代碼欄之真實姓名，轉為跨科一致假名代碼如 `STU_115_0001`）
3. **答對率與等級**（保留能力向度與等級，脫敏個人資訊）
4. **班平均彙整表**（校名替換為別名，輸出檔名自動同步脫敏）
5. **各校等級資料**（多層次合併表頭校名與備註脫敏）
6. **學校試題分析**（題目鑑別度與選答分佈 100% 完整保留）
7. **縣市試題分析**（常模參考數據與鑑別度 100% 保留）

支援多種輸出格式：
- 🚀 **全格式打包**：同時產出 Excel (`.xlsx`) + JSON (`.json`) + Markdown (`.md`)
- 📊 **Excel + JSON 雙格式**
- 📑 **Markdown 報表**：純文字輕量表格，適合貼入 AI（ChatGPT / Claude）、Notion 或 Obsidian
- 僅 Excel 活頁簿 / 僅 JSON 結構化檔案

---

## 🔒 資安與隱私合規保證

1. **Client-Side 0 上傳**：所有 Excel 解析、假名化轉換與 Zip 壓縮均於使用者的本機記憶體中運算，絕無後端伺服器、絕不上傳任何個資至雲端。
2. **公開儲存庫保護**：專案透過 `.gitignore` 嚴格排除任何原始名冊與 Excel 檔案，確保 GitHub 儲存庫乾淨安全。
3. **完全開源透明**：所有核心處理邏輯均公開可稽核。
