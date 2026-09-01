"""Excel 管理級數整理工具。

此程式需在 Windows 且已安裝 Microsoft Excel 的環境執行，透過 pywin32 讀取
Excel 原生 COM 屬性 Interior.ColorIndex，以保留格式、公式、巨集與工作表設定。
"""

from __future__ import annotations

import importlib.util
import os
import platform
import queue
import re
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


DEFAULT_RANGE = "M1:R2000"
MANAGEMENT_HEADER = "管理級數"
LEVEL_HEADERS = ["第一級", "第二級", "第三級", "第四級"]
COLOR_TO_LEVEL = {-4142: "第一級", 34: "第二級", 6: "第三級", 3: "第四級"}
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
SAVE_FORMATS = {".xlsx": 51, ".xlsm": 52, ".xls": 56}
PYWIN32_INSTALL_MESSAGE = (
    "找不到 pywin32 套件（No module named 'win32com'）。\n"
    "請在 Windows 的同一個 Python 環境安裝：pip install pywin32\n"
    "若使用 VS Code，請確認已選擇安裝 pywin32 的 Python Interpreter。"
)
WINDOWS_EXCEL_MESSAGE = "本工具需在 Windows 且已安裝 Microsoft Excel 的環境執行，才能使用 Excel COM 讀取 ColorIndex。"


class ExcelProcessError(Exception):
    """可顯示給使用者的 Excel 處理錯誤。"""


@dataclass(frozen=True)
class ExcelRange:
    """使用者輸入的 Excel 範圍。"""

    start_col: int
    start_row: int
    end_col: int
    end_row: int

    @property
    def row_count(self) -> int:
        return self.end_row - self.start_row + 1

    @property
    def data_row_count(self) -> int:
        return max(0, self.end_row - self.start_row)


@dataclass
class ProcessStats:
    """整理完成後回傳給 GUI 的統計資訊。"""

    worksheet_name: str = ""
    actual_range: str = ""
    scanned_rows: int = 0
    non_empty_cells: int = 0
    level_counts: dict[str, int] = field(default_factory=lambda: {name: 0 for name in LEVEL_HEADERS})
    ignored_color_cells: int = 0
    output_path: str = ""

    def to_message(self) -> str:
        return (
            f"實際處理的工作表名稱：{self.worksheet_name}\n"
            f"實際處理的範圍：{self.actual_range}\n"
            f"掃描的資料列數：{self.scanned_rows}\n"
            f"掃描的非空白儲存格數：{self.non_empty_cells}\n"
            f"第一級寫入項目數：{self.level_counts['第一級']}\n"
            f"第二級寫入項目數：{self.level_counts['第二級']}\n"
            f"第三級寫入項目數：{self.level_counts['第三級']}\n"
            f"第四級寫入項目數：{self.level_counts['第四級']}\n"
            f"忽略的其他 ColorIndex 儲存格數：{self.ignored_color_cells}\n"
            f"輸出檔案路徑：{self.output_path}"
        )


def column_letter_to_index(column_letters: str) -> int:
    """將欄名（例如 M、AA）轉為 1-based 欄號。"""
    value = 0
    for char in column_letters.upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def column_index_to_letter(column_index: int) -> str:
    """將 1-based 欄號轉為 Excel 欄名。"""
    letters = ""
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def validate_range(range_text: str) -> ExcelRange:
    """驗證並解析 A1:B2 形式的資料範圍。"""
    match = re.fullmatch(r"\s*([A-Za-z]{1,3})(\d+)\s*:\s*([A-Za-z]{1,3})(\d+)\s*", range_text or "")
    if not match:
        raise ExcelProcessError("資料範圍格式錯誤，請使用類似 M1:R2000 的格式。")

    start_col = column_letter_to_index(match.group(1))
    start_row = int(match.group(2))
    end_col = column_letter_to_index(match.group(3))
    end_row = int(match.group(4))

    if start_col > end_col or start_row > end_row:
        raise ExcelProcessError("資料範圍格式錯誤，請使用類似 M1:R2000 的格式。")
    if start_col < 1 or end_col > EXCEL_MAX_COLUMNS or start_row < 1 or end_row > EXCEL_MAX_ROWS:
        raise ExcelProcessError("指定範圍超出 Excel 可使用範圍。")
    if start_row != 1:
        raise ExcelProcessError("範圍起始列必須為第 1 列，因為第一列為來源表頭。")
    return ExcelRange(start_col, start_row, end_col, end_row)


def format_range(excel_range: ExcelRange) -> str:
    """將 ExcelRange 轉回 A1:B2 字串。"""
    return f"{column_index_to_letter(excel_range.start_col)}{excel_range.start_row}:{column_index_to_letter(excel_range.end_col)}{excel_range.end_row}"


def is_blank(value: Any) -> bool:
    """判斷 None、空字串、只有空白字元的字串是否為空白。"""
    return value is None or (isinstance(value, str) and value.strip() == "")


def normalize_color_index(color_index: Any) -> int | None:
    """將 COM 回傳的 ColorIndex 轉為整數，無法轉換時回傳 None。"""
    try:
        return int(color_index)
    except (TypeError, ValueError):
        return None


def default_output_path(input_path: str) -> str:
    """依原始副檔名產生預設輸出路徑，xlsm 會維持 xlsm。"""
    path = Path(input_path)
    extension = path.suffix.lower()
    output_extension = extension if extension in {".xlsm", ".xls"} else ".xlsx"
    return str(path.with_name(f"{path.stem}_整理完成{output_extension}"))


class ExcelProcessor:
    """封裝所有 Excel COM 操作，避免 GUI 與資料處理邏輯混在一起。"""

    def __init__(self, progress_callback=None, log_callback=None) -> None:
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.excel = None
        self.workbook = None
        self.com_initialized = False

    def log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)

    def update_progress(self, value: int, maximum: int) -> None:
        if self.progress_callback:
            self.progress_callback(value, maximum)

    def load_worksheet_names(self, file_path: str) -> list[str]:
        """以獨立 Excel 執行個體讀取工作表名稱。"""
        self._validate_input_file(file_path)
        self._open_excel(file_path, read_only=True)
        try:
            return [sheet.Name for sheet in self.workbook.Worksheets]
        finally:
            self.release_excel_objects(save_changes=False)

    def process_excel_data(self, file_path: str, worksheet_name: str, range_text: str, output_path: str) -> ProcessStats:
        """主要整理流程。"""
        original_range = validate_range(range_text)
        self._validate_paths(file_path, output_path)
        self._open_excel(file_path, read_only=False)
        try:
            worksheet = self._get_worksheet(worksheet_name)
            self._ensure_writable(output_path)
            source_headers = self._capture_source_headers(worksheet, original_range)
            management_col = self.find_management_level_column(worksheet)
            level_columns, inserted_count = self.create_level_columns(worksheet, management_col)
            adjusted_range, adjusted_headers = self._adjust_source_range(original_range, source_headers, management_col, inserted_count)
            self._validate_adjusted_headers(adjusted_headers)
            self._clear_old_results(worksheet, level_columns, adjusted_range.end_row)
            stats = self._scan_and_write(worksheet, worksheet_name, adjusted_range, adjusted_headers, level_columns)
            stats.output_path = output_path
            self.save_output_file(output_path)
            return stats
        finally:
            self.release_excel_objects(save_changes=False)

    def find_management_level_column(self, worksheet) -> int:
        """在第一列尋找「管理級數」表頭。"""
        last_col = worksheet.Cells(1, worksheet.Columns.Count).End(-4159).Column  # xlToLeft
        for col in range(1, last_col + 1):
            if str(worksheet.Cells(1, col).Value or "").strip() == MANAGEMENT_HEADER:
                return col
        raise ExcelProcessError("找不到表頭「管理級數」，請確認 Excel 第一列內容。")

    def create_level_columns(self, worksheet, management_col: int) -> tuple[dict[str, int], int]:
        """建立或重用第一級至第四級欄位。"""
        existing = self._find_level_columns_after_management(worksheet, management_col)
        if all(header in existing for header in LEVEL_HEADERS):
            self.log("已存在第一級至第四級表頭，將直接重用既有欄位。")
            return {header: existing[header] for header in LEVEL_HEADERS}, 0

        insert_at = management_col + 1
        worksheet.Columns(f"{column_index_to_letter(insert_at)}:{column_index_to_letter(insert_at + 3)}").Insert(Shift=1)
        for offset, header in enumerate(LEVEL_HEADERS):
            target_col = insert_at + offset
            worksheet.Cells(1, management_col).Copy()
            worksheet.Cells(1, target_col).PasteSpecial(Paste=-4122)  # xlPasteFormats
            worksheet.Cells(1, target_col).Value = header
            worksheet.Columns(management_col).Copy()
            worksheet.Columns(target_col).PasteSpecial(Paste=-4122)
        worksheet.Application.CutCopyMode = False
        self.log("已在管理級數後方新增第一級至第四級欄位。")
        return {header: insert_at + index for index, header in enumerate(LEVEL_HEADERS)}, 4

    def get_cell_color_index(self, cell) -> int | None:
        """取得 Excel 原生 Interior.ColorIndex。"""
        return normalize_color_index(cell.Interior.ColorIndex)

    def save_output_file(self, output_path: str) -> None:
        """依副檔名另存新檔，保留 xlsm 巨集格式。"""
        extension = Path(output_path).suffix.lower()
        file_format = SAVE_FORMATS.get(extension, 51)
        self.workbook.SaveAs(os.path.abspath(output_path), FileFormat=file_format)

    def release_excel_objects(self, save_changes: bool = False) -> None:
        """關閉本程式建立的 Workbook 與 Excel Application。"""
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=save_changes)
            self.workbook = None
        if self.excel is not None:
            self.excel.Quit()
            self.excel = None
        if self.com_initialized:
            import pythoncom

            pythoncom.CoUninitialize()
            self.com_initialized = False

    def _ensure_excel_com_available(self) -> None:
        """檢查 Windows、Excel COM 與 pywin32 是否可用，並提供清楚錯誤訊息。"""
        if platform.system() != "Windows":
            raise ExcelProcessError(WINDOWS_EXCEL_MESSAGE)
        if importlib.util.find_spec("win32com") is None or importlib.util.find_spec("pythoncom") is None:
            raise ExcelProcessError(PYWIN32_INSTALL_MESSAGE)

    def _open_excel(self, file_path: str, read_only: bool) -> None:
        self._ensure_excel_com_available()

        import pythoncom
        import win32com.client

        # 背景執行緒使用 COM 前需初始化，避免 tkinter 執行緒與 Excel COM 互相干擾。
        pythoncom.CoInitialize()
        self.com_initialized = True
        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            self.workbook = self.excel.Workbooks.Open(os.path.abspath(file_path), ReadOnly=read_only)
        except Exception as exc:
            self.release_excel_objects(save_changes=False)
            raise ExcelProcessError(f"無法啟動或開啟 Microsoft Excel，請確認已安裝 Excel 且檔案未被鎖定：{exc}") from exc

    def _validate_input_file(self, file_path: str) -> None:
        if not file_path or not Path(file_path).exists():
            raise ExcelProcessError("Excel 檔案不存在，請重新選擇檔案。")
        if Path(file_path).suffix.lower() not in SAVE_FORMATS:
            raise ExcelProcessError("僅支援 .xlsx、.xlsm、.xls 檔案。")

    def _validate_paths(self, file_path: str, output_path: str) -> None:
        self._validate_input_file(file_path)
        if not output_path:
            raise ExcelProcessError("請選擇輸出檔案路徑。")
        output_parent = Path(output_path).expanduser().resolve().parent
        if not output_parent.exists():
            raise ExcelProcessError("輸出路徑的資料夾不存在。")
        if Path(output_path).suffix.lower() not in SAVE_FORMATS:
            raise ExcelProcessError("輸出檔案副檔名需為 .xlsx、.xlsm 或 .xls。")

    def _ensure_writable(self, output_path: str) -> None:
        """以暫存檔測試輸出目錄寫入權限。"""
        directory = Path(output_path).resolve().parent
        try:
            with tempfile.NamedTemporaryFile(dir=directory, delete=True):
                pass
        except OSError as exc:
            raise ExcelProcessError(f"輸出路徑無法寫入：{exc}") from exc

    def _get_worksheet(self, worksheet_name: str):
        for sheet in self.workbook.Worksheets:
            if sheet.Name == worksheet_name:
                return sheet
        raise ExcelProcessError("工作表不存在，請重新選擇工作表。")

    def _capture_source_headers(self, worksheet, excel_range: ExcelRange) -> list[tuple[int, str]]:
        headers = []
        for col in range(excel_range.start_col, excel_range.end_col + 1):
            value = worksheet.Cells(excel_range.start_row, col).Value
            if is_blank(value):
                raise ExcelProcessError("來源範圍的表頭不可為空白。")
            headers.append((col, str(value).strip()))
        return headers

    def _find_level_columns_after_management(self, worksheet, management_col: int) -> dict[str, int]:
        found = {}
        for offset in range(1, 5):
            col = management_col + offset
            header = str(worksheet.Cells(1, col).Value or "").strip()
            if header in LEVEL_HEADERS:
                found[header] = col
        return found

    def _adjust_source_range(self, original_range: ExcelRange, headers: list[tuple[int, str]], management_col: int, inserted_count: int) -> tuple[ExcelRange, list[tuple[int, str]]]:
        if inserted_count == 0:
            return original_range, headers
        insert_at = management_col + 1
        adjusted_headers = []
        for original_col, header in headers:
            adjusted_col = original_col + inserted_count if original_col >= insert_at else original_col
            adjusted_headers.append((adjusted_col, header))
        adjusted_start = original_range.start_col + inserted_count if original_range.start_col >= insert_at else original_range.start_col
        adjusted_end = original_range.end_col + inserted_count if original_range.end_col >= insert_at else original_range.end_col
        return ExcelRange(adjusted_start, original_range.start_row, adjusted_end, original_range.end_row), adjusted_headers

    def _validate_adjusted_headers(self, headers: list[tuple[int, str]]) -> None:
        for _col, header in headers:
            if header in LEVEL_HEADERS:
                raise ExcelProcessError("來源範圍包含第一級至第四級結果欄，請調整資料範圍後再執行。")

    def _clear_old_results(self, worksheet, level_columns: dict[str, int], end_row: int) -> None:
        for col in level_columns.values():
            if end_row >= 2:
                worksheet.Range(worksheet.Cells(2, col), worksheet.Cells(end_row, col)).ClearContents()

    def _scan_and_write(self, worksheet, worksheet_name: str, excel_range: ExcelRange, headers: list[tuple[int, str]], level_columns: dict[str, int]) -> ProcessStats:
        stats = ProcessStats(worksheet_name=worksheet_name, actual_range=format_range(excel_range), scanned_rows=excel_range.data_row_count)
        maximum = max(1, excel_range.data_row_count)
        for row_index, row in enumerate(range(2, excel_range.end_row + 1), start=1):
            row_results = {header: [] for header in LEVEL_HEADERS}
            for col, source_header in headers:
                cell = worksheet.Cells(row, col)
                if is_blank(cell.Value):
                    continue
                stats.non_empty_cells += 1
                color_index = self.get_cell_color_index(cell)
                level_header = COLOR_TO_LEVEL.get(color_index)
                if level_header is None:
                    stats.ignored_color_cells += 1
                    continue
                row_results[level_header].append(source_header)
                stats.level_counts[level_header] += 1
            for level_header, names in row_results.items():
                if names:
                    worksheet.Cells(row, level_columns[level_header]).Value = "".join(f"{name}；" for name in names)
            if row_index % 10 == 0 or row_index == maximum:
                self.update_progress(row_index, maximum)
        return stats


class ExcelOrganizerApp(tk.Tk):
    """tkinter GUI 主視窗。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 管理級數整理工具")
        self.geometry("850x620")
        self.message_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.file_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.range_var = tk.StringVar(value=DEFAULT_RANGE)
        self.worksheet_var = tk.StringVar()
        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 6}
        ttk.Button(self, text="選擇 Excel 檔案", command=self.select_excel_file).grid(row=0, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.file_path_var, width=90).grid(row=0, column=1, columnspan=2, sticky="ew", **padding)
        ttk.Label(self, text="工作表：").grid(row=1, column=0, sticky="w", **padding)
        self.sheet_combo = ttk.Combobox(self, textvariable=self.worksheet_var, state="readonly", width=40)
        self.sheet_combo.grid(row=1, column=1, sticky="w", **padding)
        ttk.Label(self, text="資料範圍：").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.range_var, width=40).grid(row=2, column=1, sticky="w", **padding)
        ttk.Button(self, text="選擇輸出檔案", command=self.select_output_file).grid(row=3, column=0, sticky="w", **padding)
        ttk.Entry(self, textvariable=self.output_path_var, width=90).grid(row=3, column=1, columnspan=2, sticky="ew", **padding)
        self.start_button = ttk.Button(self, text="開始整理", command=self.start_processing)
        self.start_button.grid(row=4, column=0, sticky="w", **padding)
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.grid(row=4, column=1, columnspan=2, sticky="ew", **padding)
        ttk.Label(self, text="狀態 / 執行紀錄：").grid(row=5, column=0, sticky="nw", **padding)
        self.log_text = tk.Text(self, height=22, wrap="word")
        self.log_text.grid(row=5, column=1, columnspan=2, sticky="nsew", **padding)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)

    def select_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if not file_path:
            return
        self.file_path_var.set(file_path)
        self.output_path_var.set(default_output_path(file_path))
        self._append_log("正在載入工作表名稱...")
        threading.Thread(target=self.load_worksheet_names, args=(file_path,), daemon=True).start()

    def load_worksheet_names(self, file_path: str) -> None:
        try:
            names = ExcelProcessor().load_worksheet_names(file_path)
            self.message_queue.put(("sheets", names))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))

    def select_output_file(self) -> None:
        initial = self.output_path_var.get() or default_output_path(self.file_path_var.get()) if self.file_path_var.get() else "整理完成.xlsx"
        file_path = filedialog.asksaveasfilename(title="選擇輸出檔案", initialfile=Path(initial).name, initialdir=str(Path(initial).parent), defaultextension=Path(initial).suffix or ".xlsx", filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")])
        if file_path:
            self.output_path_var.set(file_path)

    def start_processing(self) -> None:
        if os.path.abspath(self.file_path_var.get() or "") == os.path.abspath(self.output_path_var.get() or ""):
            if not messagebox.askyesno("覆蓋確認", "輸出路徑與原始檔相同，確定要覆蓋嗎？"):
                return
        self.start_button.config(state="disabled")
        self.progress.config(value=0, maximum=1)
        self._append_log("開始整理...")
        args = (self.file_path_var.get(), self.worksheet_var.get(), self.range_var.get(), self.output_path_var.get())
        threading.Thread(target=self._worker_process, args=args, daemon=True).start()

    def _worker_process(self, file_path: str, worksheet_name: str, range_text: str, output_path: str) -> None:
        processor = ExcelProcessor(progress_callback=lambda value, maximum: self.message_queue.put(("progress", (value, maximum))), log_callback=lambda message: self.message_queue.put(("log", message)))
        try:
            stats = processor.process_excel_data(file_path, worksheet_name, range_text, output_path)
            self.message_queue.put(("done", stats))
        except Exception as exc:
            self.message_queue.put(("error", str(exc)))

    def _poll_queue(self) -> None:
        while not self.message_queue.empty():
            kind, payload = self.message_queue.get()
            if kind == "sheets":
                self.sheet_combo["values"] = payload
                if payload:
                    self.worksheet_var.set(payload[0])
                self._append_log("工作表名稱載入完成。")
            elif kind == "progress":
                value, maximum = payload
                self.progress.config(maximum=maximum, value=value)
            elif kind == "log":
                self._append_log(payload)
            elif kind == "done":
                self.start_button.config(state="normal")
                self.progress.config(value=self.progress["maximum"])
                message = payload.to_message()
                self._append_log(message)
                messagebox.showinfo("完成", "整理完成！\n\n" + message)
            elif kind == "error":
                self.start_button.config(state="normal")
                self.show_error_message(payload)
        self.after(100, self._poll_queue)

    def _append_log(self, message: str) -> None:
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def show_error_message(self, message: str) -> None:
        self._append_log("錯誤：" + message)
        messagebox.showerror("錯誤", message)


def main() -> None:
    """程式進入點。"""
    app = ExcelOrganizerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
