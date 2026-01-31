# 檔案內容深度搜尋工具 (File Search Tool)

![Web Version](https://img.shields.io/badge/Version-1.6.1-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Web-lightgrey)
![Deployment](https://img.shields.io/github/actions/workflow/status/Chun-Chieh-Chang/FileSearchTool/deploy.yml?branch=main)

這是一個高品質的檔案內容深度搜尋工具，支援 Excel (.xlsx, .xls)、PDF 及 Word (.docx) 檔案內容搜尋。

## 🌐 網頁版特色 (GitHub Pages)
您可以直接在瀏覽器中使用此工具，無需安裝任何環境：
- **100% 隱私保護**: 所有檔案處理皆在您的瀏覽器本地完成，不會上傳至任何伺服器。
- **支援格式**: Excel (.xlsx, .xls)、PDF (.pdf) 及 Word (.docx)。
- **進階搜尋**: 支援 `AND` (同時包含) 與 `OR` (包含任一) 關鍵字邏輯。
- **搜尋選項**: 全詞匹配、區分大小寫、檔案類型篩選。
- **現代化介面**: 採用 Glassmorphism 設計，優雅且直覺。

## 📁 專案結構 (MECE 原則)
專案採用 MECE (Mutually Exclusive, Collectively Exhaustive) 原則進行整理：

- `/web`: 網頁版應用程式源始碼 (HTML, CSS, JS, Assets)。
- `/docs`: 專案相關文件、修訂計畫及腳本。
- `/tests`: 測試用檔案與展示頁面。
- `/legacy`: 舊版或不再直接使用的設定檔 (如 .spec)。
- `/.github`: GitHub Actions 自動部署工作流。

## 🚀 部署說明
本專案已配置 **GitHub Actions** 自動部署。
當代碼推送到 `main` 分支時，會自動觸發部署至 GitHub Pages。

**主要設定檔：**
- `.github/workflows/deploy.yml`: 定義部署至 GitHub Pages 的流程。
- `package.json`: 專案定義與 npm 腳本。

## 👤 作者資訊
- **作者**: Wesley Chang
- **日期**: 2025年5月
- **描述**: 致力於開發高效、易用的辦公自動化工具。

---
&copy; 2025 Wesley Chang. Released under the MIT License.
