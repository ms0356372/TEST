from __future__ import annotations
from pathlib import Path
from typing import Any
from analyses.base_analysis import BaseAnalysis, LogCallback, ProgressCallback
from core.excel_reader import read_active_sheet
from core.excel_writer import write_analysis_workbook
from core.exceptions import AnalysisError
from core.models import FraminghamInput, FraminghamResult
from core.utils import calculate_age, safe_number, unique_output_path

OUTPUT_HEADERS=("員工編號","員工姓名","廠區","部門","性別","年齡","收縮壓","舒張壓","膽固醇","高密度脂蛋白","抽煙","糖尿病","年齡(分數)","血壓","膽固醇","高密度脂蛋白(分數)","抽煙(分數)","糖尿病(分數)","總分","總風險機率","同年齡發生率","評估十年內風險程度","相較同年齡發生率")
REQUIRED_HEADERS=("工號","姓名","廠別","部門","性別","*收縮壓","*舒張壓","*膽固醇","HDL-C","*抽菸","既往病史")
MALE_RISK={0:"3%",1:"3%",2:"4%",3:"5%",4:"7%",5:"8%",6:"10%",7:"13%",8:"16%",9:"20%",10:"25%",11:"31%",12:"37%",13:"45%"}
FEMALE_RISK={-1:"2%",0:"2%",1:"2%",2:"3%",3:"3%",4:"4%",5:"4%",6:"5%",7:"6%",8:"7%",9:"8%",10:"10%",11:"11%",12:"13%",13:"15%",14:"18%",15:"20%",16:"24%"}

def _range_score(value: float|int|None, table: tuple[tuple[int,int],...]) -> int|None:
    if value is None: return None
    for upper, score in table:
        if value <= upper: return score
    return None

def calculate_age_score(sex:str, age:int|None)->int|None:
    tables={"男":((34,-1),(39,0),(44,1),(49,2),(54,3),(59,4),(64,5),(69,6),(100,7)),"女":((34,-9),(39,-4),(44,0),(49,3),(54,6),(59,7),(100,8))}
    return _range_score(age,tables[sex]) if sex in tables and age is not None and age>=0 else None

def calculate_blood_pressure_score(sex:str,sbp:float|None,dbp:float|None)->int|None:
    if sex not in ("男","女") or sbp is None or dbp is None:return None
    if sex=="男":
        ss=3 if sbp>=160 else 2 if sbp>=140 else 1 if sbp>=130 else 0
        ds=3 if dbp>=100 else 2 if dbp>=90 else 1 if dbp>=85 else 0
    else:
        ss=3 if sbp>=160 else 2 if sbp>=140 else 0 if sbp>=120 else -3
        ds=3 if dbp>=100 else 2 if dbp>=90 else 0 if dbp>=80 else -3
    return max(ss,ds)

def calculate_cholesterol_score(sex:str,value:float|None)->int|None:
    if sex=="男": return None if value is None else (-3 if value<160 else 0 if value<200 else 1 if value<240 else 2 if value<280 else 3)
    if sex=="女": return None if value is None else (-2 if value<160 else 0 if value<200 else 1 if value<280 else 3)
    return None

def calculate_hdl_score(sex:str,value:float|None)->int|None:
    if sex=="男": return None if value is None else (2 if value<35 else 1 if value<45 else 0 if value<60 else -2)
    if sex=="女": return None if value is None else (5 if value<35 else 2 if value<45 else 1 if value<50 else 0 if value<60 else -3)
    return None

def calculate_smoking_score(smoking:str)->int|None:return 2 if smoking=="有" else 0 if smoking=="無" else None
def calculate_diabetes_score(sex:str,diabetes:str)->int|None:return (2 if sex=="男" else 4) if diabetes=="有" and sex in ("男","女") else 0 if diabetes=="無" and sex in ("男","女") else None
def calculate_total_score(*scores:int|None)->int|None:return None if any(x is None for x in scores) else sum(scores) # type: ignore[arg-type]
def calculate_total_risk(sex:str,score:int|None)->str|None:
    if score is None:return None
    if sex=="男": return "2%" if score<=-1 else "≥53%" if score>=14 else MALE_RISK.get(score)
    if sex=="女": return "1%" if score<=-2 else "≥27%" if score>=17 else FEMALE_RISK.get(score)
    return None

def calculate_age_incidence(sex:str,age:int|None)->str|None:
    if age is None or age<30 or age>100:return None
    table=((34,"2%"),(39,"3%"),(49,"4%"),(54,"6%"),(59,"7%"),(64,"9%"),(69,"11%"),(100,"14%")) if sex=="男" else ((34,"<1%"),(39,"1%"),(44,"2%"),(49,"3%"),(54,"5%"),(59,"7%"),(100,"8%")) if sex=="女" else ()
    for upper,result in table:
        if age<=upper:return result
    return None

def parse_percentage_for_comparison(value:str|None)->float|None:
    if not value:return None
    text=value.strip().replace("%","")
    try:return float(text[1:])-1e-6 if text.startswith("<") else float(text.lstrip("≥>"))
    except ValueError:return None

def calculate_risk_level(risk:str|None)->str|None:
    n=parse_percentage_for_comparison(risk)
    return None if n is None else "低度" if n<10 else "中度" if n<=20 else "高度" if n<=30 else "極高"
def compare_with_age_incidence(risk:str|None,incidence:str|None)->str|None:
    a,b=parse_percentage_for_comparison(risk),parse_percentage_for_comparison(incidence)
    if a is None or b is None:return None
    return "較低" if a<b else "一樣" if a==b else "較高"
def analyze_record(source:FraminghamInput)->FraminghamResult:
    r=FraminghamResult(source); r.age_score=calculate_age_score(source.sex,source.age); r.blood_pressure_score=calculate_blood_pressure_score(source.sex,source.systolic_bp,source.diastolic_bp); r.cholesterol_score=calculate_cholesterol_score(source.sex,source.cholesterol); r.hdl_score=calculate_hdl_score(source.sex,source.hdl); r.smoking_score=calculate_smoking_score(source.smoking); r.diabetes_score=calculate_diabetes_score(source.sex,source.diabetes); r.total_score=calculate_total_score(r.age_score,r.blood_pressure_score,r.cholesterol_score,r.hdl_score,r.smoking_score,r.diabetes_score); r.total_risk=calculate_total_risk(source.sex,r.total_score); r.age_incidence=calculate_age_incidence(source.sex,source.age); r.risk_level=calculate_risk_level(r.total_risk); r.age_comparison=compare_with_age_incidence(r.total_risk,r.age_incidence); return r

class FraminghamV1Analysis(BaseAnalysis):
    key="framingham_v1"; name="心血管風險(佛萊明罕第一版)"; required_templates=("總表",); required_headers=REQUIRED_HEADERS; output_headers=OUTPUT_HEADERS
    def run(self,templates:dict[str,str],output_dir:str,log:LogCallback=lambda _:None,progress:ProgressCallback=lambda _a,_b:None)->Path:
        path=templates.get("總表");
        if not path: raise AnalysisError("沒有選擇總表。")
        headers,rows=read_active_sheet(path); missing=[h for h in REQUIRED_HEADERS if h not in headers]
        if missing: raise AnalysisError("總表缺少必要表頭："+"、".join(missing))
        if "年齡" not in headers and "出生年月" not in headers: raise AnalysisError("總表缺少「年齡」及「出生年月」，無法計算年齡。")
        results=[]; total=len(rows); log(f"總資料筆數：{total}")
        for index,row in enumerate(rows,2):
            sex=row.get("性別"); age_num=safe_number(row.get("年齡")) if "年齡" in headers else None
            age=int(age_num) if age_num is not None and 0<=age_num<=150 else calculate_age(row.get("出生年月")) if "年齡" not in headers else None
            if age is None:log(f"第 {index} 列：年齡資料無法解析。")
            if sex not in ("男","女"):log(f"第 {index} 列：性別資料無法辨識。")
            nums=[safe_number(row.get(h)) for h in ("*收縮壓","*舒張壓","*膽固醇","HDL-C")]
            for h,n in zip(("*收縮壓","*舒張壓","*膽固醇","HDL-C"),nums):
                if n is None:log(f"第 {index} 列：{h} 資料無法轉換為數字。")
            inp=FraminghamInput(row.get("工號"),row.get("姓名"),row.get("廠別"),row.get("部門"),sex,age,*nums,"無" if row.get("*抽菸")=="從未吸菸" else "有","有" if "糖尿病" in str(row.get("既往病史") or "") else "無")
            result=analyze_record(inp); results.append(result)
            if result.total_score is None:log(f"第 {index} 列：評分資料不完整，總分及依賴欄位留空。")
            progress(len(results),total)
        target=unique_output_path(output_dir,"心血管風險_佛萊明罕第一版"); write_analysis_workbook(target,OUTPUT_HEADERS,[r.as_row() for r in results]); return target
