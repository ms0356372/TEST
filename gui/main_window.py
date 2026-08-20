from __future__ import annotations
import queue,threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog,messagebox,ttk
from config.registry import registry
from core.exceptions import AnalysisError
TEMPLATES=("總表","肌肉骨骼原稿","過負荷原稿","中高齡原稿")
class MainWindow:
 def __init__(self,root:tk.Tk)->None:
  self.root=root;root.title("Excel 健檢分析工具");root.geometry("820x720");root.minsize(720,620);self.events:queue.Queue[tuple]=queue.Queue();self.paths={k:tk.StringVar(value="尚未選擇") for k in TEMPLATES};self.selected={a.key:tk.BooleanVar() for a in registry.all()};self.output=tk.StringVar();self.progress=tk.DoubleVar();self.progress_text=tk.StringVar(value="0%");self._build();root.after(100,self._poll)
 def _build(self)->None:
  outer=ttk.Frame(self.root,padding=12);outer.pack(fill="both",expand=True);ttk.Label(outer,text="Excel 健檢分析工具",font=("TkDefaultFont",18,"bold")).pack(pady=(0,10))
  files=ttk.LabelFrame(outer,text="模板檔案",padding=8);files.pack(fill="x")
  for i,name in enumerate(TEMPLATES):ttk.Label(files,text=name,width=16).grid(row=i,column=0,sticky="w");ttk.Label(files,textvariable=self.paths[name]).grid(row=i,column=1,sticky="w",padx=8);ttk.Button(files,text="選擇／重新選擇",command=lambda n=name:self._choose(n)).grid(row=i,column=2,padx=3);ttk.Button(files,text="清除",command=lambda n=name:self.paths[n].set("尚未選擇")).grid(row=i,column=3)
  files.columnconfigure(1,weight=1);checks=ttk.LabelFrame(outer,text="分析項目",padding=8);checks.pack(fill="x",pady=8)
  for i,a in enumerate(registry.all()):ttk.Checkbutton(checks,text=a.name,variable=self.selected[a.key]).grid(row=i//2,column=i%2,sticky="w",padx=8,pady=2)
  out=ttk.LabelFrame(outer,text="輸出位置",padding=8);out.pack(fill="x");ttk.Entry(out,textvariable=self.output).pack(side="left",fill="x",expand=True);ttk.Button(out,text="選擇資料夾",command=self._choose_output).pack(side="left",padx=5);self.start=ttk.Button(out,text="開始分析",command=self._start);self.start.pack(side="left")
  status=ttk.Frame(outer);status.pack(fill="x",pady=8);ttk.Progressbar(status,variable=self.progress,maximum=100).pack(side="left",fill="x",expand=True);ttk.Label(status,textvariable=self.progress_text,width=6).pack(side="left")
  logf=ttk.LabelFrame(outer,text="執行紀錄",padding=5);logf.pack(fill="both",expand=True);self.log=tk.Text(logf,height=14,state="disabled",wrap="word");scroll=ttk.Scrollbar(logf,command=self.log.yview);self.log.configure(yscrollcommand=scroll.set);self.log.pack(side="left",fill="both",expand=True);scroll.pack(side="right",fill="y")
 def _choose(self,name:str)->None:
  p=filedialog.askopenfilename(filetypes=[("Excel 活頁簿","*.xlsx")]);
  if p:self.paths[name].set(p);self._log(f"已載入{name}：{Path(p).name}")
 def _choose_output(self)->None:
  p=filedialog.askdirectory();
  if p:self.output.set(p)
 def _start(self)->None:
  keys=[k for k,v in self.selected.items() if v.get()]
  if not keys:return messagebox.showwarning("提醒","請至少選擇一項分析。")
  if not self.output.get():return messagebox.showwarning("提醒","請選擇輸出資料夾。")
  templates={k:v.get() for k,v in self.paths.items() if v.get()!="尚未選擇"};self.start.state(["disabled"]);self.progress.set(0);threading.Thread(target=self._worker,args=(keys,templates,self.output.get()),daemon=True).start()
 def _worker(self,keys:list[str],templates:dict[str,str],output:str)->None:
  errors=[]
  for pos,key in enumerate(keys):
   a=registry.get(key);self.events.put(("log",f"開始執行：{a.name}"))
   try:
    path=a.run(templates,output,lambda m:self.events.put(("log",m)),lambda done,total:self.events.put(("progress",(pos+(done/total if total else 1))/len(keys)*100)));self.events.put(("log",f"輸出檔案：{path}"))
   except AnalysisError as exc:errors.append(f"{a.name}：{exc}");self.events.put(("log",f"錯誤：{exc}"))
   except Exception as exc:errors.append(f"{a.name}：未預期錯誤 {exc}");self.events.put(("log",f"未預期錯誤：{exc}"))
  self.events.put(("done",errors))
 def _poll(self)->None:
  try:
   while True:
    kind,value=self.events.get_nowait()
    if kind=="log":self._log(value)
    elif kind=="progress":self.progress.set(value);self.progress_text.set(f"{value:.0f}%")
    elif kind=="done":self.start.state(["!disabled"]);self.progress.set(100);self.progress_text.set("100%");messagebox.showerror("完成（含錯誤）","\n".join(value)) if value else messagebox.showinfo("完成","所有分析已完成。")
  except queue.Empty:pass
  self.root.after(100,self._poll)
 def _log(self,text:str)->None:self.log.configure(state="normal");self.log.insert("end",text+"\n");self.log.see("end");self.log.configure(state="disabled")
