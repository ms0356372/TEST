"""Tkinter GUI application."""
from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app.config import STATUS_OPTIONS
from app.db import Database
from app.services.checkin_service import CheckinError, CheckinService
from app.services.excel_importer import ExcelImporter, ImportValidationError
from app.services.query_service import QueryService


class UltrasoundCheckinApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("超音波健檢報到自動排程系統（第一階段單機版）")
        self.geometry("1200x760")
        self.minsize(1000, 700)

        self.db = Database()
        self.importer = ExcelImporter(self.db)
        self.checkin_service = CheckinService(self.db)
        self.query_service = QueryService(self.db)

        self.selected_examinee_id: int | None = None
        self.selected_checkin_record_id: int | None = None

        self._build_style()
        self._build_ui()
        self.refresh_list()

    def _build_style(self) -> None:
        style = ttk.Style(self)
        style.configure("TLabel", font=("Microsoft JhengHei", 12))
        style.configure("TButton", font=("Microsoft JhengHei", 12), padding=8)
        style.configure("Header.TLabel", font=("Microsoft JhengHei", 18, "bold"))

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        home_tab = ttk.Frame(notebook)
        import_tab = ttk.Frame(notebook)
        checkin_tab = ttk.Frame(notebook)
        list_tab = ttk.Frame(notebook)

        notebook.add(home_tab, text="首頁")
        notebook.add(import_tab, text="Excel 匯入")
        notebook.add(checkin_tab, text="報到頁")
        notebook.add(list_tab, text="名單查詢")

        self._build_home_tab(home_tab)
        self._build_import_tab(import_tab)
        self._build_checkin_tab(checkin_tab)
        self._build_list_tab(list_tab)

    def _build_home_tab(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="超音波健檢報到自動排程系統", style="Header.TLabel").pack(
            pady=20
        )
        ttk.Label(
            parent,
            text="使用流程：\n1) 先到 Excel 匯入頁載入名單\n2) 於報到頁輸入身分證進行報到\n3) 於名單查詢頁更新狀態與篩選查詢",
            justify="left",
        ).pack(anchor="w", padx=30)

    def _build_import_tab(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="匯入 Excel 名單", style="Header.TLabel").pack(anchor="w", pady=10)
        ttk.Label(
            frame,
            text="必要欄位：序號、工號(或人員工號)、姓名(或人員姓名)、性別、排程時段、項目、身分證(或身份證/ID)",
            wraplength=980,
        ).pack(anchor="w", pady=10)

        ttk.Button(frame, text="選擇並匯入 Excel", command=self.import_excel).pack(anchor="w")

    def _build_checkin_tab(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="身分證查詢與報到", style="Header.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 20)
        )

        ttk.Label(frame, text="輸入身分證字號：").grid(row=1, column=0, sticky="w")
        self.national_id_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.national_id_var, width=40).grid(
            row=1, column=1, sticky="w", padx=10
        )
        ttk.Button(frame, text="查詢", command=self.lookup_examinee).grid(row=1, column=2)

        self.checkin_info_var = tk.StringVar(value="尚未查詢")
        ttk.Label(frame, textvariable=self.checkin_info_var, justify="left").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=20
        )

        self.checkin_button = ttk.Button(
            frame, text="報到", command=self.do_checkin, state="disabled"
        )
        self.checkin_button.grid(row=3, column=0, sticky="w")

    def _build_list_tab(self, parent: ttk.Frame) -> None:
        wrapper = ttk.Frame(parent, padding=10)
        wrapper.pack(fill="both", expand=True)

        filter_frame = ttk.Frame(wrapper)
        filter_frame.pack(fill="x", pady=8)

        ttk.Label(filter_frame, text="報到篩選").pack(side="left")
        self.checkin_filter_var = tk.StringVar(value="all")
        ttk.Combobox(
            filter_frame,
            textvariable=self.checkin_filter_var,
            values=["all", "checked_in", "not_checked_in"],
            state="readonly",
            width=15,
        ).pack(side="left", padx=5)

        ttk.Label(filter_frame, text="組別篩選").pack(side="left", padx=(10, 0))
        self.group_filter_var = tk.StringVar(value="all")
        ttk.Combobox(
            filter_frame,
            textvariable=self.group_filter_var,
            values=["all", "A", "B", "C", "D", "E", "F", "G"],
            state="readonly",
            width=8,
        ).pack(side="left", padx=5)

        ttk.Label(filter_frame, text="搜尋").pack(side="left", padx=(10, 0))
        self.keyword_var = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.keyword_var, width=24).pack(side="left", padx=5)

        ttk.Button(filter_frame, text="查詢", command=self.refresh_list).pack(side="left", padx=8)

        columns = (
            "serial_no",
            "employee_id",
            "name",
            "national_id",
            "timeslot",
            "item",
            "checked",
            "checkin_no",
            "status",
        )
        self.tree = ttk.Treeview(wrapper, columns=columns, show="headings", height=20)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_list_item)

        labels = {
            "serial_no": "序號",
            "employee_id": "工號",
            "name": "姓名",
            "national_id": "身分證",
            "timeslot": "排程時段",
            "item": "項目",
            "checked": "是否報到",
            "checkin_no": "報到編號",
            "status": "目前狀態",
        }

        for key in columns:
            self.tree.heading(key, text=labels[key])
            self.tree.column(key, width=120, anchor="center")

        status_frame = ttk.Frame(wrapper)
        status_frame.pack(fill="x", pady=8)
        self.status_var = tk.StringVar(value=STATUS_OPTIONS[0])
        ttk.Label(status_frame, text="更新狀態：").pack(side="left")
        ttk.Combobox(
            status_frame,
            textvariable=self.status_var,
            values=STATUS_OPTIONS,
            state="readonly",
            width=18,
        ).pack(side="left", padx=5)
        ttk.Button(status_frame, text="套用狀態", command=self.update_selected_status).pack(
            side="left", padx=8
        )

    def import_excel(self) -> None:
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if not file_path:
            return

        try:
            result = self.importer.import_file(file_path)
            messagebox.showinfo(
                "匯入成功",
                f"新增 {result.imported_count} 筆，更新 {result.updated_count} 筆。",
            )
            self.refresh_list()
        except ImportValidationError as exc:
            messagebox.showerror("匯入失敗", str(exc))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("系統錯誤", f"匯入時發生未預期錯誤：{exc}")

    def lookup_examinee(self) -> None:
        national_id = self.national_id_var.get().strip()
        if not national_id:
            messagebox.showwarning("輸入錯誤", "請先輸入身分證字號。")
            return

        row = self.checkin_service.find_examinee(national_id)
        if not row:
            self.checkin_info_var.set("查無資料，請確認身分證或先匯入名單。")
            self.checkin_button.configure(state="disabled")
            return

        checkin_text = row["checkin_no"] if row["checkin_no"] else "尚未報到"
        info = (
            f"姓名：{row['name']}\n"
            f"工號：{row['employee_id']}\n"
            f"排程時段：{row['timeslot']}\n"
            f"項目：{row['item']}\n"
            f"報到編號：{checkin_text}"
        )
        self.checkin_info_var.set(info)
        self.checkin_button.configure(state="normal" if not row["checkin_no"] else "disabled")

    def do_checkin(self) -> None:
        national_id = self.national_id_var.get().strip()
        if not national_id:
            messagebox.showwarning("輸入錯誤", "請先輸入身分證字號。")
            return

        try:
            checkin_no = self.checkin_service.check_in(national_id)
            messagebox.showinfo("報到成功", f"報到完成，編號：{checkin_no}")
            self.lookup_examinee()
            self.refresh_list()
        except CheckinError as exc:
            messagebox.showwarning("報到失敗", str(exc))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("系統錯誤", f"報到時發生未預期錯誤：{exc}")

    def refresh_list(self) -> None:
        rows = self.query_service.list_examinees(
            status_filter=self.checkin_filter_var.get() if hasattr(self, "checkin_filter_var") else "all",
            group_filter=self.group_filter_var.get() if hasattr(self, "group_filter_var") else "all",
            keyword=self.keyword_var.get() if hasattr(self, "keyword_var") else "",
        )

        for item in self.tree.get_children() if hasattr(self, "tree") else []:
            self.tree.delete(item)

        if not hasattr(self, "tree"):
            return

        for row in rows:
            checked = "已報到" if row["checkin_no"] else "未報到"
            iid = str(row["id"])
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    row["serial_no"],
                    row["employee_id"],
                    row["name"],
                    row["national_id"],
                    row["timeslot"],
                    row["item"],
                    checked,
                    row["checkin_no"] or "",
                    row["current_status"] or "",
                ),
            )

    def on_select_list_item(self, _event: object) -> None:
        selected = self.tree.selection()
        if not selected:
            self.selected_examinee_id = None
            self.selected_checkin_record_id = None
            return

        self.selected_examinee_id = int(selected[0])
        self.selected_checkin_record_id = self.query_service.get_checkin_record_id_by_examinee(
            self.selected_examinee_id
        )

    def update_selected_status(self) -> None:
        if not self.selected_checkin_record_id:
            messagebox.showwarning("操作失敗", "請先選擇已報到的人員。")
            return

        status = self.status_var.get()
        try:
            self.checkin_service.update_status(self.selected_checkin_record_id, status)
            messagebox.showinfo("狀態更新", f"狀態已更新為：{status}")
            self.refresh_list()
        except CheckinError as exc:
            messagebox.showwarning("更新失敗", str(exc))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("系統錯誤", f"更新狀態時發生未預期錯誤：{exc}")
