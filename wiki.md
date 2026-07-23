---
title: Outlook 聯絡人名單自動更新與維護指引
date: 2026-07-23
tags:
  - outlook
  - wizard-gui
  - minecraft-style
  - automation
  - mail-contacts
  - mcp
  - obsidian
aliases:
  - Outlook 聯絡人更新
  - Minecraft 聯絡人更新精靈
---

# Outlook 聯絡人名單自動更新與維護指引

本頁面記錄了 Outlook 2016 聯絡人自動更新安裝精靈 (Minecraft 農村草地像素風格)、實機測試結果 (覆蓋 / 共存 / 比對 三大模式)、備份與還原機制、MCP Server 架設指引，以及 Obsidian Skills 部署記錄。

> [!success] 最新版本執行與實機測試結果 (2026-07-23)
> - **發行版本**：`Outlook_Contact_Updater_Wizard.exe` (v2.0 單一獨立執行檔，15.2 MB)
> - **測試範例檔案**：`C:\Users\richard_zhang\Downloads\260702mailaddress.rar`
> - **測試項目**：三大核心 Function 實機測試 100% 通過！
>   1. **覆蓋匯入 (OVERWRITE)**：成功完成 367 筆舊聯絡人全備份 ➔ 清空預設通訊錄 ➔ 匯入 367 筆新聯絡人（重複項: 0）。
>   2. **共存匯入 (COEXIST)**：成功完成 367 筆全備份 ➔ 保留舊聯絡人 ➔ 追加 367 筆新聯絡人（總數達 734 筆，精準比對出 356 筆重複 Email 聯絡人與 11 筆群組）。
>   3. **比對匯入 (SMART_MERGE)**：成功完成 734 筆全備份 ➔ 智慧篩選過濾 356 筆重複項目 ➔ 寫入非重複項目與群組。
> - **最終庫狀態**：經自動復原，預設聯絡人庫維持乾淨完整之 367 筆聯絡人。

---

## 專案結構

此專案腳本與相關文件均放置於 `D:\Program_Coding\Outllok_update`：

- `dist/Outlook_Contact_Updater_Wizard.exe` — **Minecraft 風格獨立安裝精靈 (雙擊即可運行)**。
- `wizard_gui.py` — Minecraft 農村草地像素風格 GUI 精靈主程式 (Tkinter + EXP Bar + Real-time Logs)。
- `mcp_outlook_core.py` — Outlook COM MAPI 核心引擎與三大升級匯入模式處理模組 (含 `safe_copy_and_move` COM 重試機制)。
- `test_three_functions.py` — 三大核心功能實機測試腳本。
- `build_exe.py` — PyInstaller 單一可執行檔打包自動化腳本。
- `tools/UnRAR.exe` — 內建可攜式解壓組件，無需安裝 WinRAR 即可直接解壓 `.rar` / `.zip` 通訊錄。
- `import_contacts.ps1` — 原始 PowerShell 自動化備份與匯入腳本。
- `wiki.md` — 本說明文件（Obsidian Note）。
- `mcp_analysis.md` — Outlook MCP Server 可行性分析報告。
- `MCP/` — MCP 伺服器程式庫：
  - `server.py` — Python FastMCP Server 主程式。
  - `test_client.py` — 本地測試連線工具。

---

## Minecraft 風格安裝精靈 (GUI Wizard) 使用說明

### 雙擊執行檔位置
`D:\Program_Coding\Outllok_update\dist\Outlook_Contact_Updater_Wizard.exe`

### 視覺與介面特色
- **Minecraft 農村草地像素視覺 (Rural Grassland Pixel Aesthetic)**：採用泥土層、草地綠標題列與像素風格黑體字型。
- **經驗值進度條 (EXP Bar)**：以 Minecraft 經典綠色經驗條 (`EXP LEVEL: X%`) 即時繪製與呈現數據處理階段。
- **環境自動偵測**：啟動時自動檢查 Outlook 是否安裝、是否正在背景執行（提供一鍵關閉 Outlook 功能）與解壓引擎狀態。

### 三大升級模式說明
1. 🌾 **覆蓋匯入 (OVERWRITE Mode)**: 完整備份舊資料到 `連絡人_Backup_YYYYMMDD_HHMMSS` ➔ 清空預設通訊錄 ➔ 匯入新聯絡人。
2. 🪵 **共存匯入 (COEXIST Mode)**: 完整備份舊資料 ➔ 保留舊資料 ➔ 直接追加新聯絡人 ➔ 產出 Email 重複項目清單。
3. ⛏️ **比對匯入 (SMART_MERGE Mode)**: 完整備份舊資料 ➔ 建立 Email/姓名 索引 ➔ 僅寫入全新聯絡人（跳過重複者）。

---

## Outlook MCP Server (LLM 控制)

為了讓 LLM 代理（如 Cursor、Claude Code、Antigravity）能夠直接檢索與維護您的 Outlook 聯絡人，我們在 `D:\Program_Coding\Outllok_update\MCP\` 部署了 FastMCP 伺服器。

### 1. 提供的工具 (Tools)
- `search_contacts(query: str)` — 依姓名/信箱模糊搜尋聯絡人。
- `get_contacts_count()` — 回傳預設聯絡人與備份資料夾內的項目數。
- `list_backups()` — 列出目前所有的備份目錄。
- `import_new_pst(pst_path: str)` — 線上執行聯絡人 PST 備份與覆蓋匯入。

### 2. 本地連線測試
執行以下命令即可測試 stdio 協定是否正常打通：
```bash
cd D:\Program_Coding\Outllok_update\MCP
python test_client.py
```

### 3. 如何配置到 Cursor / Claude Desktop

#### 方案 A：在 Cursor 中使用
1. 開啟 Cursor -> `Settings` -> `Features` -> `MCP`。
2. 點選 `+ Add New MCP Server`。
3. 填入設定值：
   - **Name**: `Outlook`
   - **Type**: `command`
   - **Command**: `python D:\Program_Coding\Outllok_update\MCP\server.py`
4. 點選保存，Status 顯示為綠燈即可在 Cursor 聊天中直接以自然語言控制 Outlook。

#### 方案 B：在 Claude Desktop 中使用
將以下 JSON 設定寫入 `%APPDATA%\Claude\claude_desktop_config.json` 中：
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

## 開發環境 Obsidian Skills 部署記錄

本開發環境中的 Cursor、Claude 和 Antigravity 平台均已成功安裝 `kepano/obsidian-skills` 擴充。

### 1. 安裝結果與路徑

| 平台 | 安裝路徑 | 狀態 |
|------|---------|------|
| **Cursor** | `C:\Users\richard_zhang\.cursor\skills\` | 已安裝（含 `references/`） |
| **Claude** | `C:\Users\richard_zhang\.claude\skills\` | 已安裝（Junction 指向 `.cursor\skills`） |
| **Antigravity** | `C:\Users\richard_zhang\.antigravity-ide\extensions\...\skills\` | 已安裝（已補齊 `references/`） |

### 2. 已啟用的 Skills 清單
- `obsidian-markdown` — 處理 Obsidian 專用 Markdown 語法（Wikilinks, Callouts, Embeds）。
- `obsidian-bases` — 支援 `.base` 檔案與 Obsidian 公式。
- `json-canvas` — 處理 Obsidian `.canvas` 畫布檔案。
- `obsidian-cli` — 透過 CLI 指令與執行中的 Obsidian 實例互動。
- `defuddle` — 將網頁內容轉換成乾淨的 Markdown 格式。

---

## 安全防護與資料還原

為了預防幻覺與誤刪，本專案提供雙重資料安全機制：

> [!important] 備份機制
> 每當執行更新或測試時，原有的所有聯絡人都會被**完整複製**到一個新建的子資料夾（如 `連絡人_Backup_20260723_130925`）。

### 手動還原步驟
1. 開啟 Outlook 2016。
2. 切換到「聯絡人」視圖。
3. 展開左側的連絡人目錄，找到相應時間戳記的備份資料夾（如 `連絡人_Backup_20260723_130925`）。
4. 點選該資料夾，使用 `Ctrl+A` 選取所有聯絡人。
5. 拖曳或使用右鍵將選取的聯絡人移動回預設的「連絡人」資料夾即可。
