class AnalysisError(Exception):
    """使用者可修正的分析錯誤。"""

class AnalysisNotConfigured(AnalysisError):
    """分析規則尚未提供。"""
