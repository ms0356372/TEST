from pathlib import Path
from analyses.base_analysis import BaseAnalysis,LogCallback,ProgressCallback
from core.exceptions import AnalysisNotConfigured
class PlaceholderAnalysis(BaseAnalysis):
    required_headers=();output_headers=()
    def run(self,templates:dict[str,str],output_dir:str,log:LogCallback=lambda _:None,progress:ProgressCallback=lambda _a,_b:None)->Path:
        raise AnalysisNotConfigured("此分析模組尚未設定分析規則。")
