from analyses.base_analysis import BaseAnalysis
from analyses.framingham_v1 import FraminghamV1Analysis
from analyses.interview_recommendation import InterviewRecommendationAnalysis
from analyses.middle_aged import MiddleAgedAnalysis
from analyses.musculoskeletal import MusculoskeletalAnalysis
from analyses.overload import OverloadAnalysis
class AnalysisRegistry:
    def __init__(self)->None:self._items:dict[str,BaseAnalysis]={}
    def register(self,analysis:BaseAnalysis)->None:self._items[analysis.key]=analysis
    def all(self)->tuple[BaseAnalysis,...]:return tuple(self._items.values())
    def get(self,key:str)->BaseAnalysis:return self._items[key]
registry=AnalysisRegistry()
for cls in (MusculoskeletalAnalysis,OverloadAnalysis,FraminghamV1Analysis,InterviewRecommendationAnalysis,MiddleAgedAnalysis):registry.register(cls())
