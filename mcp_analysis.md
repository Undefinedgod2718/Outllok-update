---
title: Outlook MCP Server 可行性分析與設計報告
date: 2026-07-02
tags:
  - mcp
  - outlook
  - architecture
  - json
  - llm-agent
---

# Outlook MCP Server 可行性分析與設計報告

本報告探討為 Outlook 2016 建立 Model Context Protocol (MCP) 伺服器的可行性，以供 LLM 代理（如 Cursor, Claude Code, Antigravity）直接透過 JSON-RPC 控制 Outlook 進行聯絡人、郵件及行事曆管理。

---

## 1. 架構概述

Model Context Protocol (MCP) 基於 **JSON-RPC 2.0** 協議，通常以標準輸入輸出（stdio）作為傳輸媒介。
Outlook 具備完善的 **COM (Component Object Model)** 介面，可藉由 C++、C#、Python (win32com) 或 Node.js (win32com) 進行腳本控制。

### 建議架構設計
我們採用 **Node.js + TypeScript**（基於 `@modelcontextprotocol/sdk`）作為 MCP Server 主體，透過 **node-win32com** 或調用精簡的 **PowerShell/JSON 橋接器** 來操作 Outlook COM。

```
+------------+               +--------------------+               +--------------+
| LLM Client |  JSON-RPC     | Outlook MCP Server |  win32com/COM  | Outlook App  |
| (Claude)   | ------------> | (Node.js/TypeScript| -------------> | (Outlook2016)|
|            |  (via stdio)  |  Process)          |                |              |
+------------+               +--------------------+               +--------------+
```

---

## 2. 核心功能與工具 (Tools) 設計

MCP Server 將向 LLM Client 暴露以下基於 JSON 格式的 Tools：

### A. 聯絡人查詢 (`outlook_search_contacts`)
- **輸入 (JSON Schema)**:
  ```json
  {
    "type": "object",
    "properties": {
      "query": { "type": "string", "description": "搜尋關鍵字，可匹配姓名或 Email" }
    },
    "required": ["query"]
  }
  ```
- **回傳**: 符合條件之聯絡人 JSON 陣列。

### B. 聯絡人建立/更新 (`outlook_upsert_contact`)
- **輸入 (JSON Schema)**:
  ```json
  {
    "type": "object",
    "properties": {
      "fullName": { "type": "string" },
      "email": { "type": "string" },
      "company": { "type": "string" },
      "categories": { "type": "array", "items": { "type": "string" } }
    },
    "required": ["fullName", "email"]
  }
  ```

### C. 聯絡人備份與還原 (`outlook_manage_backup`)
- **動作**: `create_backup` 或 `restore_backup`，自動化執行我們在 `import_contacts.ps1` 中採用的備份與清空邏輯。

---

## 3. 技術可行性分析

### 利基 (Advantages)
1. **本地執行安全性高**：因為 Outlook COM 物件僅限於 Windows 本地運行，不需對外開放 API 端口，安全隱私有保障。
2. **開發難度低**：利用微軟的 `Outlook.Application` COM 物件，可以直接在數秒內查詢與操縱上千筆聯絡人。
3. **生態適配性強**：透過標準的 MCP SDK，LLM 可以在無需人工介入的情況下自動獲取聯絡人、分析收發件人、並直接撰寫/發送草稿。

### 潛在風險與挑戰 (Risks & Mitigations)

> [!warning] 1. 安全性警告彈窗 (Security Prompts)
> - **問題**：Outlook 在外部程式存取 `Email1Address` 或發送郵件時，預設會彈出安全警告（如「有程式正在嘗試存取您的電子郵件地址」），並要求使用者手動點選確認，這會中斷 LLM 的自動化流程。
> - **對策**：
>   - 在受控制的公司環境中，可藉由調整組策略（GPO）中的「程式設計存取安全性」為「不要警告我」。
>   - 或使用第三方工具如 `Redemption` 庫繞過警告。

> [!important] 2. Windows 獨佔限制
> - **問題**：Outlook COM API 僅能在裝有 Windows 桌上型 Outlook 的環境執行。
> - **對策**：此 MCP Server 應標註為 Windows 專用。

---

## 4. 可行性評估結論

> [!success] 評估：高度可行 (Feasible with GPO adjustments)
> 製作 Outlook MCP Server 技術難度極低，核心在於處理 Outlook 的 COM 安全彈窗。一旦解除彈窗限制，LLM 將能流暢地完成查聯絡人、安排日曆、回覆郵件等高價值工作。
