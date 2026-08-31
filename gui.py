"""tkinter/ttk 主畫面與啟動檢查畫面。"""
from __future__ import annotations

import os
import queue
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from config_manager import get_app_dir, load_config, save_config
from excel_processor import get_sheet_names, parse_start_cell, preview_rows, process_excel
from logger import get_logger
from rule_manager import export_rules, import_rules
from self_check import run_self_check
from version import APP_NAME, APP_VERSION


class MainWindow:
    """一般使用者操作的單一主視窗。"""
    def __init__(self, root: tk.Tk):
        self.root = root; self.config, _ = load_config(); self.path: Path | None = None; self.events = queue.Queue()
        root.title(f"{APP_NAME} v{APP_VERSION}"); root.geometry("900x700"); root.minsize(760, 600)
        self._variables(); self._layout(); self._refresh_rules(); root.protocol("WM_DELETE_WINDOW", self.close)

    def _variables(self):
        c = self.config
        self.file_var=tk.StringVar(); self.sheet_var=tk.StringVar(value=c["last_sheet"]); self.start_var=tk.StringVar(value=c["start_cell"])
        self.in_var=tk.StringVar(value=c["input_delimiter"]); self.out_var=tk.StringVar(value=c["output_delimiter"])
        self.unknown_var=tk.StringVar(value="保留原始字串" if c["unknown_behavior"]=="keep" else "忽略該字串")
        self.dedupe_var=tk.BooleanVar(value=c["remove_duplicates"]); self.header_var=tk.StringVar(value=c["result_header"])
        self.create_header_var=tk.BooleanVar(value=c["create_header"]); self.status_var=tk.StringVar(value="請選擇 Excel 檔案")

    def _layout(self):
        root=self.root; root.columnconfigure(0,weight=1); root.rowconfigure(2,weight=1)
        a=ttk.LabelFrame(root,text="A. Excel 檔案",padding=10); a.grid(row=0,column=0,sticky="ew",padx=10,pady=6); a.columnconfigure(1,weight=1)
        ttk.Label(a,text="Excel：").grid(row=0,column=0); ttk.Entry(a,textvariable=self.file_var,state="readonly").grid(row=0,column=1,sticky="ew"); ttk.Button(a,text="瀏覽…",command=self.browse).grid(row=0,column=2,padx=4)
        ttk.Label(a,text="工作表：").grid(row=1,column=0); self.sheet_combo=ttk.Combobox(a,textvariable=self.sheet_var,state="readonly"); self.sheet_combo.grid(row=1,column=1,sticky="ew")
        line=ttk.Frame(a); line.grid(row=2,column=0,columnspan=3,sticky="ew",pady=4)
        ttk.Label(line,text="開始儲存格：").pack(side="left"); ttk.Entry(line,textvariable=self.start_var,width=10).pack(side="left")
        ttk.Checkbutton(line,text="建立結果欄標題",variable=self.create_header_var).pack(side="left",padx=12); ttk.Entry(line,textvariable=self.header_var,width=28).pack(side="left")
        b=ttk.LabelFrame(root,text="B. 字串設定",padding=10); b.grid(row=1,column=0,sticky="ew",padx=10,pady=4)
        ttk.Label(b,text="輸入分隔符：").pack(side="left"); ttk.Entry(b,textvariable=self.in_var,width=6).pack(side="left")
        ttk.Label(b,text="輸出分隔符：").pack(side="left",padx=(14,0)); ttk.Entry(b,textvariable=self.out_var,width=6).pack(side="left")
        ttk.Checkbutton(b,text="移除重複項目",variable=self.dedupe_var).pack(side="left",padx=16)
        ttk.Label(b,text="未知代碼：").pack(side="left"); ttk.Combobox(b,textvariable=self.unknown_var,values=("保留原始字串","忽略該字串"),state="readonly",width=14).pack(side="left")
        c=ttk.LabelFrame(root,text="C. 轉換規則（順序即輸出順序）",padding=8); c.grid(row=2,column=0,sticky="nsew",padx=10,pady=4); c.rowconfigure(0,weight=1); c.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(c,columns=("order","source","target"),show="headings",selectmode="browse"); [self.tree.heading(x,text=t) for x,t in zip(("order","source","target"),("順序","原始值","轉換值"))]; self.tree.column("order",width=60,stretch=False); self.tree.grid(row=0,column=0,sticky="nsew")
        buttons=ttk.Frame(c); buttons.grid(row=1,column=0,sticky="ew",pady=5)
        for text,cmd in (("新增",self.add_rule),("修改",self.edit_rule),("刪除",self.delete_rule),("上移",lambda:self.move(-1)),("下移",lambda:self.move(1)),("匯入規則",self.import_rules),("匯出規則",self.export_rules)): ttk.Button(buttons,text=text,command=cmd).pack(side="left",padx=3)
        d=ttk.LabelFrame(root,text="D. 執行",padding=8); d.grid(row=3,column=0,sticky="ew",padx=10,pady=6)
        ttk.Button(d,text="預覽",command=self.preview).pack(side="left"); self.run_btn=ttk.Button(d,text="開始轉換",command=self.start); self.run_btn.pack(side="left",padx=8)
        self.progress=ttk.Progressbar(d,mode="determinate"); self.progress.pack(side="left",fill="x",expand=True,padx=8); ttk.Label(d,textvariable=self.status_var).pack(side="left")

    def settings(self):
        parse_start_cell(self.start_var.get())
        if not self.in_var.get(): raise ValueError("輸入分隔符不可空白")
        if not self.config["rules"]: raise ValueError("請至少設定一筆轉換規則")
        return {"start_cell":self.start_var.get().upper(),"input_delimiter":self.in_var.get(),"output_delimiter":self.out_var.get(),"remove_duplicates":self.dedupe_var.get(),"unknown_behavior":"keep" if self.unknown_var.get().startswith("保留") else "ignore","result_header":self.header_var.get(),"create_header":self.create_header_var.get(),"last_sheet":self.sheet_var.get(),"rules":self.config["rules"]}

    def save(self): self.config.update(self.settings()); save_config(self.config)
    def browse(self):
        name=filedialog.askopenfilename(filetypes=[("Excel 活頁簿","*.xlsx *.xlsm")])
        if not name:return
        try:
            self.path=Path(name); names=get_sheet_names(self.path); self.file_var.set(str(self.path)); self.sheet_combo["values"]=names
            self.sheet_var.set(self.config["last_sheet"] if self.config["last_sheet"] in names else names[0]); self.status_var.set(f"已載入：{self.path.name}")
        except Exception as exc: messagebox.showerror("無法開啟 Excel",str(exc)); get_logger().exception("load workbook")
    def _refresh_rules(self,selected=None):
        self.tree.delete(*self.tree.get_children())
        for i,r in enumerate(self.config["rules"],1): self.tree.insert("", "end", iid=str(i-1), values=(i,r["source"],r["target"]))
        if selected is not None and self.tree.exists(str(selected)): self.tree.selection_set(str(selected))
    def _dialog(self,title,source="",target=""):
        source=simpledialog.askstring(title,"原始值：",initialvalue=source,parent=self.root)
        if source is None:return None
        target=simpledialog.askstring(title,"轉換值：",initialvalue=target,parent=self.root)
        return (source.strip(),target.strip()) if target is not None else None
    def add_rule(self):
        pair=self._dialog("新增規則")
        if not pair:return
        source,target=pair
        if not source or not target: messagebox.showwarning("規則錯誤","原始值與轉換值不可空白"); return
        found=next((r for r in self.config["rules"] if r["source"]==source),None)
        if found:
            if messagebox.askyesno("規則已存在",f"規則 {source} 已存在，是否修改原設定？"): found["target"]=target
        else:self.config["rules"].append({"source":source,"target":target})
        self._refresh_rules(); self.save()
    def _index(self):
        selected=self.tree.selection(); return int(selected[0]) if selected else None
    def edit_rule(self):
        i=self._index()
        if i is None:return
        old=self.config["rules"][i]; pair=self._dialog("修改規則",old["source"],old["target"])
        if not pair:return
        if any(j!=i and r["source"]==pair[0] for j,r in enumerate(self.config["rules"])): messagebox.showwarning("規則錯誤","原始值已存在"); return
        self.config["rules"][i]={"source":pair[0],"target":pair[1]}; self._refresh_rules(i); self.save()
    def delete_rule(self):
        i=self._index()
        if i is not None and messagebox.askyesno("刪除規則","確定刪除選取規則？"): self.config["rules"].pop(i); self._refresh_rules(); save_config(self.config)
    def move(self,delta):
        i=self._index()
        if i is None or not 0<=i+delta<len(self.config["rules"]):return
        self.config["rules"][i],self.config["rules"][i+delta]=self.config["rules"][i+delta],self.config["rules"][i]; self._refresh_rules(i+delta); self.save()
    def import_rules(self):
        name=filedialog.askopenfilename(filetypes=[("JSON","*.json")])
        if name:
            try:self.config["rules"]=import_rules(Path(name));self._refresh_rules();self.save()
            except Exception as exc:messagebox.showerror("匯入失敗",f"規則 JSON 格式錯誤：{exc}")
    def export_rules(self):
        name=filedialog.asksaveasfilename(defaultextension=".json",filetypes=[("JSON","*.json")],initialfile="特殊作業規則.json")
        if name:
            try:export_rules(Path(name),self.config["rules"])
            except Exception as exc:messagebox.showerror("匯出失敗",str(exc))
    def preview(self):
        try:
            if not self.path:raise ValueError("請先選擇 Excel 檔案")
            s=self.settings(); rows=preview_rows(self.path,self.sheet_var.get(),rules=s["rules"],**{k:s[k] for k in ("start_cell","input_delimiter","output_delimiter","unknown_behavior","remove_duplicates")})
            win=tk.Toplevel(self.root);win.title("轉換預覽");tree=ttk.Treeview(win,columns=("row","raw","result"),show="headings");
            for x,t in zip(("row","raw","result"),("Excel 列","原始內容","轉換結果")):tree.heading(x,text=t)
            for row,raw,result in rows:tree.insert("","end",values=(row,"" if raw is None else raw,result))
            tree.pack(fill="both",expand=True,padx=10,pady=10);win.geometry("700x420");self.save()
        except Exception as exc:messagebox.showerror("無法預覽",str(exc))
    def start(self):
        try:
            if not self.path:raise ValueError("請先選擇 Excel 檔案")
            s=self.settings(); suffix=self.path.suffix.lower(); default=f"{self.path.stem}_轉換完成_{datetime.now():%Y%m%d_%H%M%S}{suffix}"
            name=filedialog.asksaveasfilename(defaultextension=suffix,initialfile=default,filetypes=[("Excel",f"*{suffix}")])
            if not name:return
            dest=Path(name)
            if dest.resolve()==self.path.resolve():raise ValueError("禁止覆蓋原始 Excel，請選擇其他檔名")
            self.save();self.run_btn["state"]="disabled";self.status_var.set("轉換中……");self.progress["value"]=0
            threading.Thread(target=self._worker,args=(dest,s),daemon=True).start();self.root.after(80,self._poll)
        except Exception as exc:messagebox.showerror("無法開始",str(exc))
    def _worker(self,dest,s):
        try:
            result=process_excel(self.path,dest,s["last_sheet"],progress=lambda n,t:self.events.put(("progress",n,t)),**{k:s[k] for k in ("start_cell","rules","input_delimiter","output_delimiter","unknown_behavior","remove_duplicates","result_header","create_header")});self.events.put(("done",dest,result))
        except Exception as exc:get_logger().exception("conversion failed");self.events.put(("error",str(exc)))
    def _poll(self):
        try:
            while True:
                event=self.events.get_nowait()
                if event[0]=="progress":_,n,t=event;self.progress["maximum"]=max(t,1);self.progress["value"]=n;self.status_var.set(f"轉換中：{n} / {t}")
                elif event[0]=="done":_,path,r=event;self.run_btn["state"]="normal";self.status_var.set(f"完成：{r['processed']} 筆，未知代碼：{r['unknown']} 個");messagebox.showinfo("輸出成功",f"已儲存：\n{path}");return
                else:self.run_btn["state"]="normal";messagebox.showerror("轉換失敗",event[1]);return
        except queue.Empty:self.root.after(80,self._poll)
    def close(self):
        try:self.save()
        except Exception:pass
        self.root.destroy()


def launch_with_splash() -> None:
    """用 after 逐項呈現背景檢查結果，完成後才進主畫面。"""
    root=tk.Tk();root.title("程式啟動檢查");root.geometry("520x380");root.resizable(False,False)
    ttk.Label(root,text=APP_NAME,font=("Microsoft JhengHei",18,"bold")).pack(pady=(18,4));ttk.Label(root,text="系統檢查中…").pack()
    box=tk.Text(root,height=12,width=58,state="disabled");box.pack(padx=15,pady=10);bar=ttk.Progressbar(root,maximum=8);bar.pack(fill="x",padx=25);events=queue.Queue()
    def callback(item,count):events.put((item,count))
    def worker():events.put(("finished",run_self_check(callback)))
    def poll():
        try:
            while True:
                item,count=events.get_nowait()
                if item=="finished":
                    if any(x["status"]=="FAIL" for x in count):
                        messagebox.showerror("程式啟動檢查失敗","無法通過啟動檢查，請查看 logs 資料夾。")
                        ttk.Button(root,text="開啟錯誤紀錄資料夾",command=lambda:os.startfile(get_app_dir()/"logs")).pack();return
                    root.destroy();main=tk.Tk();MainWindow(main);main.mainloop();return
                symbol={"PASS":"✓","WARNING":"!","FAIL":"✕"}[item["status"]];box["state"]="normal";box.insert("end",f"[{symbol}] {item['name']}：{item['message']}\n");box.see("end");box["state"]="disabled";bar["value"]=count
        except queue.Empty:root.after(60,poll)
    threading.Thread(target=worker,daemon=True).start();root.after(60,poll);root.mainloop()
