# zjgj-Fleet-n-Plate-Collector

[繁體中文](#繁體中文) | [简体中文](#简体中文) | [English](#english)

---

## 繁體中文

### 專案簡介
`zjgj-Fleet-n-Plate-Collector` 是一個自動化公車數據收集工具。本程式能自動掃描所有指定公車線路的上行與下行資訊，獲取包含線路號、上下行方向、車輛自編號、車牌號碼、發車時間以及駕駛員工號等詳細資料，並自動匯出至 Excel 檔案。支援排程/定時執行，每日持續收集全天線路數據，新數據將自動追加至現有 Excel 檔案中。

### 主要功能
- 🚌 **全線路掃描**：支援指定公車線路的上行與下行雙向資料擷取。
- 📊 **多維度數據**：完整收集線路號、上下行、車輛自編號、車牌號、發車時間及駕駛員工號。
- ⏰ **定時自動運行**：可設定排程定時執行，實現全天候數據追蹤。
- 📁 **增量追加儲存**：新獲取的數據會自動追加儲存至 Excel 檔案（.xlsx），不覆蓋歷史紀錄。

### ⚠️ 特別聲明
> **本程式碼完全是由 AI 自動生成的垃圾屎山**。  
> 如有異議，請提交 [Issues](../../issues) 或親臨 **幹將西路 456 號** 進行線下面談。

---

## 简体中文

### 项目简介
`zjgj-Fleet-n-Plate-Collector` 是一个自动化公交数据收集工具。本程序能够自动扫描所有指定公交线路的上行与下行信息，获取包含线路号、上下行方向、车辆自编号、车牌号码、发车时间以及驾驶员工号等详细数据，并自动输出至 Excel 文件。支持定时/计划任务运行，实现全天线路信息的持续采集，新增数据会自动追加至现有的 Excel 文件中。

### 主要功能
- 🚌 **全线路扫描**：支持指定公交线路的上行与下行双向数据抓取。
- 📊 **多维度数据**：完整收集线路号、上下行、车辆自编号、车牌号、发车时间及驾驶员工号。
- ⏰ **定时自动运行**：可配置定时任务，实现全天候数据采集与追踪。
- 📁 **增量追加保存**：新获取的数据会自动追加保存至 Excel 文件（.xlsx），无需担心覆盖历史记录。

### ⚠️ 特别声明
> **本代码完全是由 AI 自动生成的垃圾屎山**。  
> 如有异议，请提交 [Issues](../../issues) 或到 **干将西路 456 号** 进行线下面谈。

---

## English

### Overview
`zjgj-Fleet-n-Plate-Collector` is an automated public transit data collection tool. It automatically scans specified bus routes (both outbound and inbound directions) to extract detailed fleet and operational information—including route numbers, direction, fleet numbers, license plate numbers, dispatch times, and driver IDs—and exports everything into Excel files. Designed for scheduled execution, it continuously captures transit data throughout the day, automatically appending newly acquired records into the Excel sheets.

### Key Features
- 🚌 **Full Route Scanning**: Scans both outbound and inbound trips for all specified bus lines.
- 📊 **Comprehensive Data Capture**: Collects route IDs, directions, fleet numbers, license plates, departure schedules, and driver IDs.
- ⏰ **Scheduled Execution**: Supports scheduled/cron runs for continuous, round-the-clock data collection.
- 📁 **Incremental Excel Logging**: Automatically appends new data records to Excel files (.xlsx) without overwriting historical data.

### ⚠️ Disclaimer
> **This codebase is entirely AI-generated trash code.**  
> If you have any objections or complaints, feel free to submit an [Issue](../../issues) or drop by **No. 456 Ganjiang West Road** for an in-person confrontation.
