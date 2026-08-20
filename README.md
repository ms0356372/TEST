# Excel 健檢問卷整理與分析工具

以 tkinter/ttk 與 openpyxl 製作的 Windows 桌面工具。可同時選擇多個分析；目前完整實作「心血管風險（佛萊明罕第一版）」與「中高齡」分析，其餘模組安全地回報規則尚未設定。本工具僅執行指定評分規則，不取代醫療專業判斷。

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

專案同時支援本機打包與 GitHub Actions 雲端打包；兩種方式皆使用 `--onefile --noconsole`，產生不顯示命令提示字元視窗的單一 EXE。

### 方法一：Windows 本機打包

1. 安裝含 `py` launcher 的 Python。
2. 雙擊專案根目錄的 `build_exe.bat`；腳本會檢查 Python，並自動安裝 `requirements.txt` 與 PyInstaller。
3. 打包完成後，成品位於 `dist/ExcelHealthTool.exe`。

### 方法二：使用 GitHub Actions 打包

Workflow 位於 `.github/workflows/build-windows-exe.yml`，會在 push 到 `main` 時自動執行，也可手動執行：

1. 開啟 GitHub repository 頁面。
2. 點選頁面上方的 **Actions**。
3. 在左側選擇 **Build Windows EXE**。
4. 點選 **Run workflow**，選擇 `main` 分支後，再按一次綠色的 **Run workflow**。
5. 等待工作流程中的依賴安裝、測試與 PyInstaller 打包步驟全部完成並顯示綠色勾號。
6. 點入已完成的 workflow run，捲動至頁面下方的 **Artifacts** 區域。
7. 點選 **Excel健檢分析工具-Windows** 下載 Artifact ZIP。
8. 解壓縮後即可取得 `Excel健檢分析工具.exe`。

GitHub Actions 使用 `windows-latest` 與 Python 3.12；只有所有測試成功後才會打包及上傳 Artifact。Artifact 預設保留 30 天。目標 Windows 電腦不需安裝 Python。

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
- 肌肉骨骼、過負荷、面談建議尚無規則，選擇後會明確提示而不當機。

## 中高齡分析

「中高齡」分析已完整實作，且只使用 GUI 匯入的 **中高齡原稿**。輸入使用第一列精確表頭、第二列起的非空資料；來源檔以唯讀方式開啟，結果另存為 `中高齡_YYYYMMDD_HHMMSS.xlsx`。

### 年齡與輸出

- 原稿有 `年齡` 表頭時直接使用該欄，不要求日期欄。
- 原稿沒有 `年齡` 時，必須同時有 `出生日期` 與 `體檢日期`；分別取前三碼民國年，以「體檢民國年 − 出生民國年」計算，不判斷生日是否已過。
- 無法解析的單筆日期或答案會記入 Log，該分數與依賴結果留空，不會中止整批。
- 結果固定輸出 A～AO 共 41 欄：A～AD 為基本資料與原始答案、AE～AK 為第 1～7 題分數、AL 為總分、AM 為等級、AN 為意義、AO 為措施宗旨。

### 評分規則

1. **第 1 題（AE）**：從第 1 題答案末端解析 0～10 的整數；H 顯示同一個解析結果，AE 直接沿用 H，不重複解析。
2. **第 2 題（AF）**：2.1 與 2.2 各依 `很好=5、好=4、普通=3、不好=2、很不好=1` 換算後相加；任一題無法辨識即留空。
3. **第 3 題（AG）**：逐一檢查 3.1～3.14 所列疾病 keyword，同一 keyword 在同一格重複只算一次。疾病數換算為 `0=7、1=5、2=4、3=3、4=2、5 個以上=1`。
4. **第 4 題（AH）**：六種工作影響答案依序換算為 6～1 分。
5. **第 5 題（AI）**：`0天=5、1~9天=4、10~24天=3、25~99天=2、100~365天=1`；答案採 substring contains 判斷。
6. **第 6 題（AJ）**：`不太可能=1、不確定=4、應該可以=7`。
7. **第 7 題（AK）**：三題各依 `總是=4、常常=3、有時=2、很少=1、從不=0` 相加；原始合計 0～3、4～6、7～9、10～12 分別換算成 1、2、3、4 分。

AL 是 AE～AK 七個分數的合計；任何必要分數缺失時，AL～AO 全部留空。有效總分的 AM～AO 分級如下：

| AL 總分 | AM 等級 | AN 意義 | AO 措施宗旨 |
|---|---|---|---|
| 7～37 | 弱 | 不能勝任工作要求 | 恢復其工作適能 |
| 38～42 | 普通 | 工作適能有待提高 | 改進其工作適能 |
| 43～46 | 良 | 能勝任所從事的工作 | 支持其工作適能 |
| 47～49 | 優 | 能很好地勝任所從事的工作 | 維持其工作適能 |

### 中高齡輸出樣式

所有有效儲存格沿用共用格式：標楷體 10、水平及垂直置中。第一列表頭依區段填色：A1～AD1 為 `F2F2F2`、AE1～AK1 為 `E2EFDA`、AL1～AO1 為 `D9E1F2`。
