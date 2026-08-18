from dataclasses import dataclass
from typing import Any

@dataclass(slots=True)
class FraminghamInput:
    employee_id: Any; employee_name: Any; site: Any; department: Any
    sex: Any; age: int | None; systolic_bp: float | None; diastolic_bp: float | None
    cholesterol: float | None; hdl: float | None; smoking: str; diabetes: str

@dataclass(slots=True)
class FraminghamResult:
    source: FraminghamInput
    age_score: int | None = None; blood_pressure_score: int | None = None
    cholesterol_score: int | None = None; hdl_score: int | None = None
    smoking_score: int | None = None; diabetes_score: int | None = None
    total_score: int | None = None; total_risk: str | None = None
    age_incidence: str | None = None; risk_level: str | None = None
    age_comparison: str | None = None
    def as_row(self) -> list[Any]:
        s=self.source
        return [s.employee_id,s.employee_name,s.site,s.department,s.sex,s.age,s.systolic_bp,s.diastolic_bp,s.cholesterol,s.hdl,s.smoking,s.diabetes,self.age_score,self.blood_pressure_score,self.cholesterol_score,self.hdl_score,self.smoking_score,self.diabetes_score,self.total_score,self.total_risk,self.age_incidence,self.risk_level,self.age_comparison]
