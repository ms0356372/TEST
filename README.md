# Excel 管理級數整理工具

## 1. 程式設計說明

本工具是可在 Windows 執行的 Python 3.10+ GUI 程式，使用 `tkinter` 建立介面，並以 `pywin32` / `win32com.client` 建立獨立 Excel Application 執行個體。程式透過 Excel COM 直接讀取 `cell.Interior.ColorIndex`，因此可依 Excel 原生 `ColorIndex` 分類，而不是只依賴 `openpyxl` 的色碼。

主要流程如下：

1. 使用者選擇 `.xlsx`、`.xlsm` 或 `.xls` 檔案。
2. 程式載入工作表清單，使用者選擇工作表與資料範圍，預設為 `M1:R2000`。
3. 開始整理時先驗證檔案、工作表、範圍、表頭、輸出路徑與寫入權限。
4. 在第一列尋找完全等於 `管理級數` 的表頭。
5. 若 `管理級數` 後方尚未存在 `第一級`、`第二級`、`第三級`、`第四級`，就在其後插入四欄並複製格式。
6. 插入欄位前會先記錄來源範圍的原始欄號與表頭；插入後若來源範圍需要位移，會自動修正實際讀取欄位。
7. 每次執行前先清空四個結果欄的資料列內容，不刪除表頭。
8. 只掃描使用者指定範圍內的非空白儲存格，並依 `ColorIndex` 寫入對應級數欄：
   - `-4142` → `第一級`
   - `34` → `第二級`
   - `6` → `第三級`
   - `3` → `第四級`
9. 同一列同級數多個項目會依來源欄位由左至右串接，每個表頭後保留全形分號 `；`。
10. 以另存新檔方式輸出；若來源是 `.xlsm`，輸出也維持 `.xlsm` 並使用 Excel COM 儲存以保留 VBA 巨集。
11. 無論成功或失敗都在 `finally` 關閉本程式建立的 Workbook 與 Excel Application，避免殘留 `EXCEL.EXE`。

## 2. 完整 Python 程式碼

完整程式碼位於 [`main.py`](main.py)，可直接執行。核心架構包含：

- `validate_range()`：驗證 `M1:R2000` 這類 A1 範圍。
- `find_management_level_column()`：尋找 `管理級數` 表頭。
- `create_level_columns()`：建立或重用四個級數欄。
- `get_cell_color_index()`：讀取 Excel 原生 `Interior.ColorIndex`。
- `process_excel_data()`：執行掃描、分類、寫入與另存新檔。
- `release_excel_objects()`：釋放 COM 物件。
- `ExcelOrganizerApp`：處理 GUI、背景執行緒與安全的 GUI 更新。

## 3. 套件安裝方式

請在 Windows、Python 3.10 以上、且已安裝 Microsoft Excel 的環境中執行：

```bash
python -m pip install --upgrade pip
pip install pywin32 pyinstaller
```

如果看到 `No module named 'win32com'`，代表目前執行程式的 Python 環境尚未安裝 `pywin32`，或 VS Code 選到不同的 Python Interpreter；請在同一個環境重新執行上方安裝指令。

本版本 GUI 使用 Python 內建 `tkinter`，不需要安裝 `customtkinter`。若你的 Python 發行版未包含 tkinter，請改安裝包含 Tcl/Tk 的官方 Python。

## 4. VS Code 執行方式

1. 使用 VS Code 開啟本資料夾。
2. 建立並啟用虛擬環境：

```bash
python -m venv .venv
.venv\Scripts\activate
```

3. 安裝套件：

```bash
pip install pywin32 pyinstaller
```

若仍出現 `No module named 'win32com'`，請先確認 VS Code 右下角或 `Python: Select Interpreter` 顯示的是 `.venv`。

4. 執行程式：

```bash
python main.py
```

5. 若 VS Code 找不到正確 Python，請執行 `Python: Select Interpreter` 並選擇 `.venv`。

## 5. PyInstaller 打包方式

基本打包指令：

```bash
pyinstaller --onefile --windowed --name Excel管理級數整理工具 main.py
```

若遇到 pywin32 hook 未自動收集的環境，可使用較保守的打包方式：

```bash
pyinstaller --onefile --windowed --name Excel管理級數整理工具 --hidden-import win32timezone --hidden-import pythoncom --hidden-import win32com --hidden-import win32com.client main.py
```

打包後執行檔會位於：

```text
dist\Excel管理級數整理工具.exe
```

注意：目標電腦仍需安裝 Microsoft Excel，因為本工具依賴 Excel COM 自動化。

## 6. 操作說明

1. 開啟程式後按「選擇 Excel 檔案」。
2. 確認檔案路徑、工作表下拉選單已正確載入。
3. 輸入資料範圍，例如預設 `M1:R2000`；範圍起始列必須是第 1 列。
4. 選擇輸出檔案路徑。預設會產生 `原檔名_整理完成.xlsx`；若來源是 `.xlsm`，預設輸出為 `.xlsm`。
5. 按「開始整理」。
6. 進度條與紀錄區會顯示處理狀態。
7. 完成後會跳出成功訊息，並顯示工作表、實際範圍、掃描列數、非空白儲存格數、各級寫入項目數、忽略色彩數與輸出路徑。
8. 若輸出路徑與原始檔相同，程式會先跳出覆蓋確認。

## 7. 測試方式與測試案例

建議在 Windows Excel 中準備測試活頁簿，第一列包含 `管理級數` 與來源範圍表頭，再用 Excel 填色設定指定 `ColorIndex`。

### 測試一：單一第一級

- 在某列只有一個非空白儲存格。
- 該儲存格 `ColorIndex` 為 `-4142`。
- 執行後應在同列 `第一級` 寫入該欄表頭加 `；`。

### 測試二：多個相同級數

- 同一列有三個非空白儲存格的 `ColorIndex` 都為 `34`。
- 執行後應依來源欄位由左至右，將三個表頭串接寫入 `第二級`。

### 測試三：四種級數同時存在

- 同一列分別放入 `-4142`、`34`、`6`、`3`。
- 執行後應分別寫入 `第一級`、`第二級`、`第三級`、`第四級`。

### 測試四：空白儲存格

- 將空白儲存格加上背景色。
- 執行後該空白儲存格必須被跳過，不得寫入結果。

### 測試五：其他顏色

- 將非空白儲存格設定成不屬於 `-4142`、`34`、`6`、`3` 的 `ColorIndex`。
- 執行後不得寫入四個級數欄，且忽略統計應增加。

### 測試六：重複執行

- 對同一輸出檔重複執行。
- 不可重複新增四個表頭，結果也不可重複累加。

### 測試七：來源範圍限制

- 指定 `M1:R2000`。
- 在 M 欄以前放入有顏色且有資料的儲存格。
- 執行後不得判斷或寫入 M 欄以前的儲存格。

### 測試八：插入欄位造成位移

- 將 `管理級數` 放在來源範圍之前。
- 執行新增四欄後，仍須正確讀取原本指定的來源欄位表頭與資料。

可先執行靜態語法檢查：

```bash
python -m py_compile main.py
```
