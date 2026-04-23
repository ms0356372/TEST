# 超音波健檢報到自動排程系統（第一階段：Windows 單機版）

## 第一部分：技術架構說明
- **GUI 框架：Tkinter (ttk)**
  - Python 內建，不需額外安裝大型 GUI runtime。
  - 在 Windows 上穩定、部署簡單、維護成本低。
  - 本案需求偏流程操作與資料列表，Tkinter 足夠且可靠。
- **資料庫：SQLite**
  - 單機資料儲存最佳選擇，檔案式 DB、免伺服器。
  - 支援交易（transaction），可確保報到編號不重複。
- **Excel 匯入：pandas + openpyxl**
  - 可穩定讀取 `.xlsx`。
  - 以欄位同義字自動辨識，缺欄時明確回報。
- **架構模式：模組化三層**
  - `gui`：視覺與互動
  - `services`：業務邏輯（匯入、報到、查詢）
  - `db`：資料存取與 schema

## 第二部分：專案目錄結構
```text
TEST/
├─ app/
│  ├─ main.py
│  ├─ config.py
│  ├─ db.py
│  ├─ gui/
│  │  └─ app.py
│  ├─ services/
│  │  ├─ excel_importer.py
│  │  ├─ checkin_service.py
│  │  └─ query_service.py
│  └─ utils/
│     └─ time_slot.py
├─ data/                         # SQLite 會自動建立在此
├─ requirements.txt
├─ sample_excel_template.csv
└─ README.md
```

## 第三部分：完整程式碼（逐檔）
請直接查看以下檔案（皆為可執行完整檔案）：
- `app/main.py`
- `app/config.py`
- `app/db.py`
- `app/gui/app.py`
- `app/services/excel_importer.py`
- `app/services/checkin_service.py`
- `app/services/query_service.py`
- `app/utils/time_slot.py`

## 第四部分：requirements.txt
```txt
pandas==2.3.0
openpyxl==3.1.5
pyinstaller==6.15.0
```

## 第五部分：範例 Excel 格式
已提供：`sample_excel_template.csv`（可用 Excel 開啟再另存 `.xlsx`）。

必要欄位同義字：
- 序號：`序號`
- 工號：`人員工號` / `工號`
- 姓名：`人員姓名` / `姓名`
- 性別：`性別`
- 排程時段：`排程時段`
- 項目：`項目`
- 身分證：`身分證` / `身份證` / `ID`

## 第六部分：Windows 安裝與執行說明
1. 安裝 Python 3.11+（安裝時勾選 Add Python to PATH）。
2. 開啟 PowerShell，切換到專案目錄。
3. 建立虛擬環境：
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
4. 安裝套件：
   ```powershell
   pip install -r requirements.txt
   ```
5. 啟動系統：
   ```powershell
   python -m app.main
   ```

## 第七部分：exe 打包方式
在專案根目錄執行：
```powershell
pyinstaller --noconfirm --onefile --windowed --name UltrasoundCheckin app/main.py
```
輸出檔案位置：`dist/UltrasoundCheckin.exe`

若要附帶 `sample_excel_template.csv` 可另外複製到同資料夾。

## 第八部分：未來升級成多平板同步版的保留設計
本版已預留擴充方向：
1. **Service 層隔離**：GUI 不直接碰 SQL，便於替換為 API 呼叫。
2. **流水號獨立表 `group_sequences`**：未來可改為伺服器端交易鎖控。
3. **狀態歷程 `status_history`**：可支援同步衝突解決與稽核。
4. **系統設定表 `system_settings`**：可加入診間、裝置 ID、同步旗標。
5. **欄位正規化邏輯集中**：未來可抽成共享 schema/驗證服務。

---

## 目前功能對照（第一階段必做）
- [x] Excel 匯入（含欄位同義字辨識）
- [x] 身分證查詢
- [x] 報到編號產生（依時段分組、各組流水、重啟不歸零）
- [x] 狀態更新（即時寫入 DB + 歷程）
- [x] 名單查詢、篩選（已報到/未報到、組別、關鍵字）
- [x] 本地資料永久保存（SQLite）
