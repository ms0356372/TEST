# Excel整理分析工具

Python 3.13 相容的繁體中文 Windows 桌面程式，以 tkinter/ttk 與 openpyxl 完成「匯入 Excel → 設定有序規則 → 預覽 → 插入結果欄 → 另存 Excel」。不使用 pandas，也不需要 Microsoft Excel。

## 使用方式

```bash
python -m pip install -r requirements.txt
python main.py
```

選擇 `.xlsx`/`.xlsm`、工作表及開始儲存格，調整規則後可先預覽。正式轉換會在來源欄右側**插入**新欄，其他欄右移，絕不允許覆蓋來源檔；`.xlsm` 以 `keep_vba=True` 載入並維持副檔名。

## 設定與記錄

設定位於 `%APPDATA%\ExcelTransformTool\config.json`，記錄位於其 `logs` 子資料夾。非 Windows 平台測試時使用 `~/.config/ExcelTransformTool`。設定損壞會先改名為含時間戳的 `config_corrupted_*.json`，再建立預設設定。

## 測試

```bash
python test_transform.py
python main.py --self-test
```

`test_transform.py` 包含 14 項字串及實際 Excel 插欄整合測試。Self Test 檢查版本、設定、openpyxl、排序、未知值、去重、寫入權限及中文 Excel 讀寫。

如需需求文件中的人工驗證檔，請執行：

```bash
python generate_test_excel.py
```

這會在專案目錄產生 `test_input.xlsx`。Excel 活頁簿本身是 ZIP 格式的二進位檔，部分 PR 系統不支援二進位附件，因此 Git 僅保存可重現的產生器，產生的檔案已列入 `.gitignore`。

## Windows 單檔 EXE

在 Windows 雙擊 `build_exe.bat`。腳本會檢查 Python/pip、安裝 requirements、執行單元測試與 Python Self Test、清理舊建置、以 `--onefile --windowed --clean` 打包、執行真正 EXE 的 Self Test，只有通過後才複製至：

`release\Excel整理分析工具_v1.0.0.exe`

完整過程與 EXE 大小記錄在 `build_log.txt`。正常啟動 EXE 時會先即時顯示啟動檢查，再開啟主畫面。

## 模組

- `main.py`：GUI / `--self-test` 入口。
- `gui.py`：啟動檢查、主畫面、預覽、背景轉換與進度。
- `excel_processor.py`：字串演算法、唯讀預覽、Excel 插欄與樣式複製。
- `config_manager.py`：APPDATA 設定及損壞修復。
- `rule_manager.py`：有序規則驗證、JSON 匯入匯出。
- `self_check.py`、`logger.py`、`version.py`：健康檢查、日誌、版本。
- `generate_test_excel.py`：重建人工驗證用 `test_input.xlsx`，避免 PR 夾帶二進位檔。
