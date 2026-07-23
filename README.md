# Outlook 聯絡人自動更新精靈 (Minecraft Style v2.1.1)

[![Python Version](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Release](https://img.shields.io/badge/version-v2.1.1-gold.svg)](CHANGELOG.md)

本專案提供一套具備 **Minecraft 農村草地像素風格 (Minecraft Rural Grassland Pixel Art)** 的 Outlook 聯絡人自動更新安裝精靈 GUI 與核心 MAPI 引擎。支援備份現有聯絡人、自動解壓 `.rar` / `.zip` / `.pst` 檔案，並提供覆蓋、共存、智慧比對三種升級模式。

---

## 🌾 特色亮點 (Features)

- **🖼️ 自訂品牌圖示與標誌 (Custom Branding)**：
  - Windows Executable 使用專屬檔圖示 `logo.ico`。
  - 精靈標題列嵌入獨特 UI Logo 圖章 (`unnamed.jpg`)。
- **🌾 Minecraft 視覺與 8-Bit 音效體驗**：
  - 泥土層與草地綠面板設計、黑體像素字型與 **綠色經驗值進度條 (EXP Bar)**。
  - 內建 8-bit 音效 (`click.wav`, `xp_gain.wav`, `victory.wav`) 提示互動操作。
- **📦 可攜式解壓組件**：
  - 內建可攜式 `UnRAR.exe`，使用者電腦無需安裝 WinRAR 即可直接讀取與解壓 `.rar` / `.zip` 通訊錄檔案。
- **🛡️ 雙重資料安全機制**：
  - 每次升級前自動建立 `連絡人_Backup_YYYYMMDD_HHMMSS` 資料夾備份全數原有聯絡人。
- **⚡ 三大升級匯入模式**：
  1. **覆蓋匯入 (OVERWRITE)**：備份舊資料 ➔ 清空預設通訊錄 ➔ 匯入新資料。
  2. **共存匯入 (COEXIST)**：備份舊資料 ➔ 保留舊資料 ➔ 追加新資料 ➔ 產出 Email 重複細節。
  3. **比對匯入 (SMART_MERGE)**：備份舊資料 ➔ 智慧比對索引 ➔ 僅寫入全新聯絡人（跳過重複者）。
- **🤖 FastMCP LLM 整合**：
  - 提供基於 `FastMCP` 的 Python MCP 伺服器，支援 Cursor 與 Claude Desktop 直接透過 LLM 對話查詢與維護 Outlook 聯絡人。

---

## 🚀 快速開始 (Quick Start)

### 安裝依賴與執行 GUI
```bash
# 1. 安裝套件依賴
pip install -r requirements.txt

# 2. 執行 GUI 精靈
python wizard_gui.py
```

### 執行單一獨立 .exe
雙擊執行 `dist/Outlook_Contact_Updater_Wizard.exe` 即可（無需安裝 Python 或 WinRAR）。

---

## 🧪 實機測試數據 (Test Benchmark)

使用標準通訊錄檔案 (`260702mailaddress.rar`，含 367 筆聯絡人/群組) 完成 100% 實機驗證：

| 測試模式 | 操作流程 | 實機測試數據 |
| :--- | :--- | :--- |
| **覆蓋匯入 (OVERWRITE)** | 全備份 ➔ 清空 ➔ 寫入 | 備份 367 筆 ➔ 清空 ➔ 寫入 367 筆 (重複: 0) |
| **共存匯入 (COEXIST)** | 全備份 ➔ 保留 ➔ 追加 | 備份 367 筆 ➔ 追加 367 筆 (總數 734 筆，精準偵測 356 筆 Email 重複項) |
| **比對匯入 (SMART_MERGE)** | 全備份 ➔ 索引比對 ➔ 過濾 | 備份 734 筆 ➔ 智慧過濾 356 筆重複項 ➔ 寫入非重複項目與群組 |

---

## 🤖 Outlook MCP Server (LLM 控制配置)

本專案提供 FastMCP 伺服器 (`MCP/server.py`)，可於 Cursor 或 Claude Desktop 中配置：

### Cursor 設定
1. 開啟 `Settings` ➔ `Features` ➔ `MCP` ➔ `+ Add New MCP Server`。
2. 設定內容：
   - **Name**: `Outlook`
   - **Type**: `command`
   - **Command**: `python D:\Program_Coding\Outllok_update\MCP\server.py`

### Claude Desktop 設定 (`claude_desktop_config.json`)
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "D:\\Program_Coding\\Outllok_update\\MCP\\server.py"
      ]
    }
  }
}
```

---

## 🛠️ 開發與打包

```bash
# 1. 執行實機測試
python test_three_functions.py

# 2. 生成 8-bit 音效檔
python create_sounds.py

# 3. 打包單一 EXE (含 logo.ico)
python build_exe.py
```

---

## 📜 授權與版本 (License & Versioning)

- **Current Version**: `v2.1.1` (Semantic Versioning)
- **Changelog**: 詳見 [CHANGELOG.md](CHANGELOG.md)
