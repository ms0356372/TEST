"""中高齡工作能力問卷分析。"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Mapping, Sequence

from analyses.base_analysis import BaseAnalysis, LogCallback, ProgressCallback
from core.excel_reader import read_active_sheet
from core.excel_writer import write_analysis_workbook
from core.exceptions import AnalysisError
from core.utils import safe_number, unique_output_path
from core.validator import validate_excel_input, validate_output_directory

Q1 = "1.假設您的工作能力在最好的狀況為10分，您給目前工作能力打幾分?"
Q21_SOURCE = "2.1依您目前工作所需要的體力需求來衡量您的工作能力"
Q22_SOURCE = "2.2依您目前工作所需要的心力需求來衡量您的工作能力"
Q21_OUTPUT = Q21_SOURCE + "(您目前的工作內容會有一定的基本體力要求，例如：搬運、遞送、駕駛、行走…之類，以此為基準來看您的工作能力)"
Q22_OUTPUT = Q22_SOURCE + "(您目前的工作內容會有一定的基本心力/腦力要求，例如：思考、記憶、創新、溝通、社交互動…之類，以此為基準來看您的工作能力)"
DISEASE_HEADERS = (
    "3.1因事故導致的傷害", "3.2肌肉骨骼疾病", "3.3心血管疾病", "3.4呼吸系統疾病",
    "3.5心理健康問題", "3.6神經系統和感覺器官疾病", "3.7消化系統疾病",
    "3.8生殖泌尿器官疾病", "3.9皮膚疾病", "3.10 腫瘤", "3.11內分泌或代謝疾病",
    "3.12血液疾病", "3.13先天缺陷", "3.14其它問題或疾病",
)
Q4 = "4.您的疾病或傷害，對您工作的影響為何?"
Q5 = "5.最近12個月，您曾經因為健康問題(生病、治療或醫療檢查)整天請假(或無法工作) 共多少天?"
Q6 = "6.以您的健康狀況衡量，您目前的工作還可以繼續從事兩年嗎?"
Q71 = "7.1 您最近都能輕鬆地從事日常活動嗎?"
Q72 = "7.2您最近很有活力且有警覺性嗎?"
Q73 = "7.3您最近對未來充滿希望嗎?"

OUTPUT_HEADERS = (
    "工號", "姓名", "性別", "年齡", "廠別", "部門", "課別", Q1, Q21_OUTPUT, Q22_OUTPUT,
    *DISEASE_HEADERS, Q4, Q5, Q6, Q71, Q72, Q73,
    "1(分數)", "2(分數)", "3(分數)", "4(分數)", "5(分數)", "6(分數)", "7(分數)",
    "總分", "等級", "意義", "措施宗旨",
)
REQUIRED_HEADERS = (
    "工號", "姓名", "性別", "廠別", "部門", "課別", Q1,
    *DISEASE_HEADERS, Q4, Q5, Q6, Q71, Q72, Q73,
)

QUESTION2_SCORE_MAP = {"很好": 5, "好": 4, "普通": 3, "不好": 2, "很不好": 1}
DISEASE_KEYWORDS = {
    DISEASE_HEADERS[0]: ("背部", "手臂或手部", "腿或腳", "身體其他部位"),
    DISEASE_HEADERS[1]: ("上背或頸椎的問題", "下背部的問題，重複發生的疼痛", "從背部傳到腿部的疼痛", "肌肉骨骼問題影響到四肢", "類風濕性關節炎", "其它肌肉骨骼問題"),
    DISEASE_HEADERS[2]: ("高血壓", "冠狀動脈心臟病", "冠狀動脈血栓", "心臟功能不全", "其它心血管疾病"),
    DISEASE_HEADERS[3]: ("反複的呼吸道感染", "慢性支氣管炎", "慢性鼻竇炎", "支氣管性氣喘", "肺氣腫", "肺結核", "其它呼吸系統疾病"),
    DISEASE_HEADERS[4]: ("精神疾病或嚴重心理健康問題", "輕微心理疾病或問題"),
    DISEASE_HEADERS[5]: ("聽覺問題或傷害", "視覺疾病或傷害", "神經系統疾病", "其他神經系統和感覺器官疾病"),
    DISEASE_HEADERS[6]: ("膽結石或膽囊疾病", "肝臟或胰臟及疾病", "胃潰瘍或十二指腸潰瘍", "胃炎或十二指腸不適", "大腸激躁", "其他消化系統疾病"),
    DISEASE_HEADERS[7]: ("尿道感染", "腎臟疾病", "生殖系統疾病", "其他生殖泌尿系統疾病"),
    DISEASE_HEADERS[8]: ("過敏性皮疹或紅斑", "其他疹子", "其他皮膚疾病"),
    DISEASE_HEADERS[9]: ("良性腫瘤", "惡性腫瘤"),
    DISEASE_HEADERS[10]: ("肥胖", "糖尿病", "甲狀腺腫大或其他甲狀腺疾病", "其它內分泌或代謝疾病"),
    DISEASE_HEADERS[11]: ("貧血", "其他血液問題"),
    DISEASE_HEADERS[12]: ("請寫出所有經醫師診治的先天缺陷診斷名稱",),
    DISEASE_HEADERS[13]: ("請寫出所有經醫師診治的其他問題或疾病",),
}
QUESTION3_SCORE_MAP = {0: 7, 1: 5, 2: 4, 3: 3, 4: 2}
QUESTION4_SCORE_MAP = {
    "沒有任何影響/我沒有任何疾病": 6, "會引起一些症狀，但可以工作": 5,
    "有時候需放慢工作節奏或改變工作方式": 4, "經常需放慢工作節奏或改變工作方式": 3,
    "因為疾病，我覺得只能做兼職的工作": 2, "自己覺得完全不能工作": 1,
}
QUESTION5_SCORE_MAP = {"0天": 5, "1-9天": 4, "10-24天": 3, "25-99天": 2, "100-365天": 1}
QUESTION6_SCORE_MAP = {"不太可能": 1, "不確定": 4, "應該可以": 7}
QUESTION7_RAW_SCORE_MAP = {"總是": 4, "常常": 3, "有時": 2, "很少": 1, "從不": 0}
LEVELS = (
    (37, "弱", "不能勝任工作要求", "恢復其工作適能"),
    (42, "普通", "工作適能有待提高", "改進其工作適能"),
    (46, "良", "能勝任所從事的工作", "支持其工作適能"),
    (49, "優", "能很好地勝任所從事的工作", "維持其工作適能"),
)


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _excel_column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def find_unique_header_containing(
    headers: Sequence[str], keyword: str, log: LogCallback = lambda _: None,
) -> str:
    """Return the sole header containing keyword, never guessing duplicates."""
    matches = [(index, header) for index, header in enumerate(headers, 1) if keyword in header]
    if not matches:
        raise AnalysisError(f"找不到包含「{keyword}」的表頭。")
    if len(matches) > 1:
        log(f"偵測到多個包含「{keyword}」的表頭：")
        for index, header in matches:
            log(f"{_excel_column_name(index)}欄：{header}")
        raise AnalysisError(
            "偵測到多個包含以下文字的表頭：\n\n"
            f"「{keyword}」\n\n"
            "為避免使用錯誤欄位，請確認原始 Excel 表頭。"
        )
    return matches[0][1]


def calculate_roc_age(birth_date: Any, exam_date: Any) -> int | None:
    birth, exam = _text(birth_date), _text(exam_date)
    if len(birth) < 3 or len(exam) < 3 or not birth[:3].isdigit() or not exam[:3].isdigit():
        return None
    age = int(exam[:3]) - int(birth[:3])
    return age if 0 <= age <= 150 else None


def extract_question1_score(value: Any) -> int | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
    else:
        match = re.search(r"(?:分數[:：]\s*)?(\d+(?:\.\d+)?)\s*$", _text(value))
        if not match:
            return None
        number = float(match.group(1))
    return int(number) if number.is_integer() and 0 <= number <= 10 else None


def calculate_question2_score(answer_i: Any, answer_j: Any) -> int | None:
    first, second = QUESTION2_SCORE_MAP.get(_text(answer_i)), QUESTION2_SCORE_MAP.get(_text(answer_j))
    return first + second if first is not None and second is not None else None


def extract_disease_matchable_segments(value: Any) -> list[str]:
    """Return 「、」 items with full-width-colon annotations removed.

    Each item starts a fresh matching scope, so text following a colon is
    ignored only until the next ideographic comma.
    """
    return [
        matchable
        for item in _text(value).split("、")
        if (matchable := item.split("：", 1)[0].strip())
    ]


def count_disease_items(answers: Mapping[str, Any] | Sequence[Any]) -> int:
    values = answers if isinstance(answers, Mapping) else dict(zip(DISEASE_HEADERS, answers))
    count = 0
    for header, keywords in DISEASE_KEYWORDS.items():
        segments = extract_disease_matchable_segments(values.get(header))
        count += sum(any(keyword in segment for segment in segments) for keyword in keywords)
    return count


def calculate_question3_score(disease_count: int) -> int | None:
    if disease_count < 0:
        return None
    return 1 if disease_count >= 5 else QUESTION3_SCORE_MAP.get(disease_count)


def calculate_question4_score(answer: Any) -> int | None:
    return QUESTION4_SCORE_MAP.get(_text(answer))


def calculate_question5_score(answer: Any) -> int | None:
    return QUESTION5_SCORE_MAP.get(_text(answer))


def calculate_question6_score(answer: Any) -> int | None:
    return QUESTION6_SCORE_MAP.get(_text(answer))


def calculate_question7_score(answer_ab: Any, answer_ac: Any, answer_ad: Any) -> int | None:
    scores = [QUESTION7_RAW_SCORE_MAP.get(_text(answer)) for answer in (answer_ab, answer_ac, answer_ad)]
    if any(score is None for score in scores):
        return None
    raw_total = sum(scores)  # type: ignore[arg-type]
    return 1 if raw_total <= 3 else 2 if raw_total <= 6 else 3 if raw_total <= 9 else 4


def calculate_total_score(*scores: int | None) -> int | None:
    return None if any(score is None for score in scores) else sum(scores)  # type: ignore[arg-type]


def _classification(total_score: int | None, position: int) -> str | None:
    if total_score is None or not 7 <= total_score <= 49:
        return None
    return next(item[position] for item in LEVELS if total_score <= item[0])


def calculate_work_ability_level(total_score: int | None) -> str | None:
    return _classification(total_score, 1)


def calculate_work_ability_meaning(total_score: int | None) -> str | None:
    return _classification(total_score, 2)


def calculate_work_ability_measure(total_score: int | None) -> str | None:
    return _classification(total_score, 3)


def _log_invalid(log: LogCallback, row_number: int, question: str, value: Any) -> None:
    log(f"第 {row_number} 列：第 {question} 題答案無法辨識：{_text(value)}")


def _build_output_row(
    row: Mapping[str, Any], row_number: int, use_direct_age: bool,
    question_21_header: str, question_22_header: str, log: LogCallback,
) -> list[Any]:
    if use_direct_age:
        age_number = safe_number(row.get("年齡"))
        age = int(age_number) if age_number is not None and age_number.is_integer() and 0 <= age_number <= 150 else None
        if age is None:
            log(f"第 {row_number} 列：年齡資料無法辨識：{_text(row.get('年齡'))}")
    else:
        age = calculate_roc_age(row.get("出生日期"), row.get("體檢日期"))
        if age is None:
            log(f"第 {row_number} 列：出生日期或體檢日期無法解析，年齡無法計算。")

    q1 = extract_question1_score(row.get(Q1))
    q2 = calculate_question2_score(row.get(question_21_header), row.get(question_22_header))
    disease_answers = {header: row.get(header) for header in DISEASE_HEADERS}
    q3 = calculate_question3_score(count_disease_items(disease_answers))
    q4 = calculate_question4_score(row.get(Q4))
    q5 = calculate_question5_score(row.get(Q5))
    q6 = calculate_question6_score(row.get(Q6))
    q7 = calculate_question7_score(row.get(Q71), row.get(Q72), row.get(Q73))
    for question, score, value in (("1", q1, row.get(Q1)), ("2", q2, f"{_text(row.get(question_21_header))} / {_text(row.get(question_22_header))}"), ("4", q4, row.get(Q4)), ("5", q5, row.get(Q5)), ("6", q6, row.get(Q6)), ("7", q7, f"{_text(row.get(Q71))} / {_text(row.get(Q72))} / {_text(row.get(Q73))}")):
        if score is None:
            _log_invalid(log, row_number, question, value)
    total = calculate_total_score(q1, q2, q3, q4, q5, q6, q7)
    if total is None:
        log(f"第 {row_number} 列：評分資料不完整，總分及分級欄位留空。")
    raw = [row.get(key) for key in ("工號", "姓名", "性別")]
    raw += [age]
    raw += [row.get(key) for key in ("廠別", "部門", "課別")]
    raw += [q1, row.get(question_21_header), row.get(question_22_header)]
    raw += [row.get(header) for header in DISEASE_HEADERS]
    raw += [row.get(key) for key in (Q4, Q5, Q6, Q71, Q72, Q73)]
    return raw + [q1, q2, q3, q4, q5, q6, q7, total, calculate_work_ability_level(total), calculate_work_ability_meaning(total), calculate_work_ability_measure(total)]


class MiddleAgedAnalysis(BaseAnalysis):
    key = "middle_aged"
    name = "中高齡"
    required_templates = ("中高齡原稿",)
    required_headers = REQUIRED_HEADERS
    output_headers = OUTPUT_HEADERS

    def run(self, templates: dict[str, str], output_dir: str, log: LogCallback = lambda _: None, progress: ProgressCallback = lambda _a, _b: None) -> Path:
        source = templates.get("中高齡原稿")
        if not source:
            raise AnalysisError("尚未匯入「中高齡原稿」，無法執行中高齡分析。")
        source_path = validate_excel_input(source, "中高齡原稿")
        output_path = validate_output_directory(output_dir)
        headers, rows = read_active_sheet(source_path)
        question_21_header = find_unique_header_containing(headers, Q21_SOURCE, log)
        question_22_header = find_unique_header_containing(headers, Q22_SOURCE, log)
        missing = [header for header in REQUIRED_HEADERS if header not in headers]
        if "年齡" not in headers:
            missing.extend(header for header in ("出生日期", "體檢日期") if header not in headers)
        if missing:
            raise AnalysisError("缺少表頭：\n" + "\n".join(missing))
        log(f"總資料筆數：{len(rows)}")
        output_rows = []
        for row_number, row in enumerate(rows, 2):
            output_rows.append(_build_output_row(
                row, row_number, "年齡" in headers,
                question_21_header, question_22_header, log,
            ))
            progress(len(output_rows), len(rows))
        target = unique_output_path(output_path, "中高齡")
        return write_analysis_workbook(
            target, OUTPUT_HEADERS, output_rows,
            header_fill_ranges=((1, 30, "F2F2F2"), (31, 37, "E2EFDA"), (38, 41, "D9E1F2")),
        )
