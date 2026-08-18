# Excel 健檢問卷整理與分析工具

以 tkinter/ttk 與 openpyxl 製作的 Windows 桌面工具。可同時選擇多個分析；目前完整實作使用指定規則的「心血管風險（佛萊明罕第一版）」，其餘模組安全地回報規則尚未設定。本工具僅執行指定評分規則，不取代醫療專業判斷。

## 安裝與執行

需求為 Python 3.10+（Windows 通常隨 Python 提供 tkinter）。

```bash
python -m pip install -r requirements.txt
python main.py
```

在 GUI 選擇「總表」、一或多個分析及可寫入的輸出資料夾，再按「開始分析」。工作在背景執行，GUI 以 queue 與 `after()` 安全更新紀錄及進度。輸入檔一律唯讀，輸出另存為含時間戳的 `.xlsx`。

## 測試

```bash
python -m unittest discover -s tests -v
python -m compileall -q .
```

測試涵蓋男女評分邊界、百分比比較、資料容錯、完整 Excel 流程、23 欄順序、樣式、色彩及原檔未變更。

## 打包 Windows EXE

先安裝 `python -m pip install pyinstaller`，再雙擊 `build_exe.bat`。腳本使用 `--onefile --noconsole`，成品位於 `dist/Excel健檢分析工具.exe`，目標電腦不需 Python。

## 專案架構

- `main.py`, `gui/`：入口、檔案/資料夾選擇、checkbox、背景工作、ProgressBar 與 Log。
- `core/`：唯讀 Excel reader、統一 writer/樣式、資料模型、例外與容錯工具。
- `analyses/`：抽象介面、Framingham 實作及 placeholder。
- `config/registry.py`：Analysis Registry；GUI 不含分析清單硬編碼。
- `tests/`：單元與整合測試。

## 模板與擴充

目前模板鍵為「總表、肌肉骨骼原稿、過負荷原稿、中高齡原稿」。輸入有多工作表時使用 **Active Worksheet**；reader 已保留 `sheet_name` 參數供日後 sheet 選擇器使用。新增模板時在 GUI 的 `TEMPLATES` 增加顯示項目；新增分析時繼承 `BaseAnalysis`，宣告 `key/name/required_templates/required_headers/output_headers/run()`，再於 registry 註冊，GUI 會自動顯示。

## Framingham 輸入與輸出

必要表頭採去除前後空白後的**完全相符**判斷：`工號`、`姓名`、`廠別`、`部門`、`性別`、`*收縮壓`、`*舒張壓`、`*膽固醇`、`HDL-C`、`*抽菸`、`既往病史`，以及 `年齡` 或 `出生年月` 至少一個。若有 `年齡` 欄即優先使用且不回退至出生年月。性別只接受 `男`/`女`；抽菸只有完全等於 `從未吸菸` 判為無；病史文字含 `糖尿病` 判為有。

輸出固定 A～W 23 欄。有效範圍字型為標楷體 10、水平/垂直置中，表頭 `F2F2F2`；V 欄中度/高度/極高為 `D6DCE4`/`FFFF00`/`FF0000`，W 欄較高為 `FFFF00`，低度、較低與一樣不額外填色。欄寬上限 40。

## 常見錯誤

- 缺少總表或必要表頭：該分析停止，GUI 顯示錯誤；不使用模糊欄名。
- Excel 損毀、被鎖定或輸出資料夾不可寫：確認檔案及權限後重試。
- `未檢`、`N/A`、`-` 或非數字：該列基本資料仍輸出，無法計算欄位留空並記錄，不中止整批。
- 出生年月只讀前四字作西元年，以目前年份相減，不判斷生日是否已過。
- 肌肉骨骼、過負荷、面談建議、中高齡尚無規則，選擇後會明確提示而不當機。
