# SAAA 學生學習能力檢測去識別化工具 (SAAA De-identification Tool)

[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live%20Demo-brightgreen?style=flat&logo=github)](https://liangchengyu-ntcu.github.io/SAAA-Deidentification-Tool/)
[![CI / CD](https://github.com/liangchengyu-ntcu/SAAA-Deidentification-Tool/actions/workflows/deploy.yml/badge.svg)](https://github.com/liangchengyu-ntcu/SAAA-Deidentification-Tool/actions)
[![Bootstrap 5](https://img.shields.io/badge/Bootstrap-5.3.3-purple?logo=bootstrap)](https://getbootstrap.com/)
[![SweetAlert2](https://img.shields.io/badge/SweetAlert2-Popup-red)](https://sweetalert2.github.io/)
[![100% Client-Side](https://img.shields.io/badge/Security-100%25%20Offline%20%26%20Client--Side-blue?logo=shield)](https://github.com/)

> 🔒 **100% 本機瀏覽器離線運算** · **零安裝** · **零個資上傳** · **支援新竹縣與各縣市學力檢測多結構 Excel**

---

## 📁 開源儲存庫架構 (Repository Structure)

本儲存庫為純前端輕量化架構，已透過 `.gitignore` 嚴格過濾原始資料與本地測試檔案，僅包含部署所需之純靜態程式碼：

```
📁 SAAA-Deidentification-Tool/
│
├── 📁 SAAA去識別化工具/             # ⭐️ 核心網頁工具（GitHub Pages 發布目錄）
│   ├── 📄 index.html              # HTML5 語意化首頁 (Bootstrap 5 + SweetAlert2)
│   ├── 📄 README.md               # 網頁工具說明文件
│   ├── 📁 css/                    # 本地樣式檔 (bootstrap, sweetalert2, custom)
│   ├── 📁 js/                     # 核心業務邏輯模組 (app.js)
│   └── 📁 lib/                    # 開源引擎 (bootstrap.bundle, sweetalert2, xlsx, jszip)
│
├── 📁 .github/workflows/          # ⚙️ GitHub Actions CI / CD 自動化流程
│   ├── 📄 deploy.yml              # 自動發布至 GitHub Pages
│   └── 📄 ci.yml                  # 語法檢驗與 Release 自動打包
│
├── 📄 .gitignore                  # 🛡️ 隱私過濾規則 (自動排除所有原始資料與產出)
└── 📄 README.md                   # 專案總覽文件
```

---

## 🚀 快速開始

### 方式一：直接在 GitHub Pages 線上使用 (推薦)
點擊 [GitHub Pages 線上版連結](https://liangchengyu-ntcu.github.io/SAAA-Deidentification-Tool/)，直接在瀏覽器拖入評量資料夾或 Zip，即可於本機記憶體完成脫敏並下載。

### 方式二：下載離線工具包單機使用
1. 至 [Releases 頁面](https://github.com/liangchengyu-ntcu/SAAA-Deidentification-Tool/releases) 下載 `SAAA_Deidentification_Tool_Offline.zip`。
2. 解壓縮後，**點兩下打開 `index.html`**，即使拔掉網路線亦可正常使用。

---

## 📊 多格式匯出支援 (Export Formats)

工具支援多種輸出格式切換：
- 🚀 **全格式打包**：同時產出 Excel (`.xlsx`) + JSON (`.json`) + Markdown (`.md`)
- 📊 **Excel + JSON 雙格式**：試算表與結構化資料一次搞定
- 📑 **Markdown 報表**：純文字輕量表格，適合貼入 AI（ChatGPT / Claude）、Notion 或 Obsidian
- 僅 Excel 活頁簿 / 僅 JSON 結構化檔案

---

## 🔒 資安與隱私合規保證

1. **Client-Side 0 上傳**：所有 Excel 解析、假名化轉換與 Zip 壓縮均於使用者的電腦瀏覽器記憶體中運算，絕無後端伺服器、絕不上傳任何個資至雲端。
2. **公開儲存庫保護**：專案透過 `.gitignore` 嚴格排除任何原始名冊與 Excel 檔案，確保 GitHub 儲存庫乾淨安全。
