import tempfile
import unittest
from analyses.middle_aged import (
    DISEASE_HEADERS, DISEASE_KEYWORDS, OUTPUT_HEADERS, calculate_question2_score,
    calculate_question3_score, calculate_question4_score, calculate_question5_score,
    calculate_question6_score, calculate_question7_score, calculate_roc_age,
    calculate_total_score, calculate_work_ability_level, calculate_work_ability_meaning,
    calculate_work_ability_measure, count_disease_items, extract_question1_score,
    MiddleAgedAnalysis,
)
from core.exceptions import AnalysisError

class MiddleAgedScoringTests(unittest.TestCase):
    def test_output_has_41_columns(self):
        self.assertEqual(len(OUTPUT_HEADERS), 41)

    def test_missing_template_error(self):
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaisesRegex(AnalysisError, "尚未匯入「中高齡原稿」"):
                MiddleAgedAnalysis().run({}, folder)

    def test_roc_age(self):
        self.assertEqual(calculate_roc_age("0870101", "1150819"), 28)
        self.assertEqual(calculate_roc_age("087/01/01", "115/08/19"), 28)
        for birth, exam in ((None, "1150819"), ("087", ""), ("ab7", "115"), ("08", "115"), ("087", "11"), ("115", "087")):
            with self.subTest(birth=birth, exam=exam): self.assertIsNone(calculate_roc_age(birth, exam))

    def test_question1(self):
        prefix = "0分表示目前完全無法工作，10分表示目前工作能力最佳)請選擇最適合的分數:"
        for score in (0, 5, 10):
            self.assertEqual(extract_question1_score(prefix + str(score)), score)
            self.assertEqual(extract_question1_score(score), score)
        for invalid in (None, "", prefix + "11", "無法辨識", 3.5):
            self.assertIsNone(extract_question1_score(invalid))

    def test_question2(self):
        cases = [("很好", "很好", 10), ("很好", "好", 9), ("普通", "普通", 6), ("很不好", "很不好", 2)]
        for first, second, expected in cases:
            self.assertEqual(calculate_question2_score(first, second), expected)
        self.assertIsNone(calculate_question2_score("", "很好"))

    def test_every_disease_keyword_group_and_no_duplicate_count(self):
        for header, keywords in DISEASE_KEYWORDS.items():
            with self.subTest(header=header):
                self.assertEqual(count_disease_items({header: "；".join(keywords)}), len(keywords))
        self.assertEqual(count_disease_items({DISEASE_HEADERS[0]: "背部；背部"}), 1)
        self.assertEqual(count_disease_items({DISEASE_HEADERS[0]: "背部；手臂或手部；腿或腳"}), 3)
        self.assertEqual(count_disease_items({}), 0)

    def test_question3_boundaries(self):
        expected = {0: 7, 1: 5, 2: 4, 3: 3, 4: 2, 5: 1, 6: 1, 10: 1}
        for count, score in expected.items():
            self.assertEqual(calculate_question3_score(count), score)
        self.assertIsNone(calculate_question3_score(-1))

    def test_question4_all_answers(self):
        expected = {"沒有任何影響/我沒有任何疾病":6,"會引起一些症狀，但可以工作":5,"有時候需放慢工作節奏或改變工作方式":4,"經常需放慢工作節奏或改變工作方式":3,"因為疾病，我覺得只能做兼職的工作":2,"自己覺得完全不能工作":1}
        for answer, score in expected.items(): self.assertEqual(calculate_question4_score(answer), score)
        self.assertIsNone(calculate_question4_score(""))

    def test_questions5_and6(self):
        for answer, score in {"0天":5,"1-9天":4,"10-24天":3,"25-99天":2,"100-365天":1}.items(): self.assertEqual(calculate_question5_score(answer), score)
        for answer, score in {"不太可能":1,"不確定":4,"應該可以":7}.items(): self.assertEqual(calculate_question6_score(answer), score)
        self.assertIsNone(calculate_question5_score("366天")); self.assertIsNone(calculate_question6_score("可能"))

    def test_question7_raw_boundaries(self):
        # Answers yielding raw totals 0,3,4,6,7,9,10,12.
        cases = [("從不","從不","從不",1),("常常","從不","從不",1),("總是","從不","從不",2),("常常","常常","從不",2),("總是","常常","從不",3),("常常","常常","常常",3),("總是","常常","常常",4),("總是","總是","總是",4)]
        for a,b,c,score in cases: self.assertEqual(calculate_question7_score(a,b,c), score)
        self.assertIsNone(calculate_question7_score("", "總是", "總是"))

    def test_total_and_classification_boundaries(self):
        self.assertEqual(calculate_total_score(10,10,7,6,5,7,4), 49)
        self.assertEqual(calculate_total_score(0,2,1,1,1,1,1), 7)
        self.assertIsNone(calculate_total_score(10,None,7,6,5,7,4))
        expected = {7:("弱","不能勝任工作要求","恢復其工作適能"),37:("弱","不能勝任工作要求","恢復其工作適能"),38:("普通","工作適能有待提高","改進其工作適能"),42:("普通","工作適能有待提高","改進其工作適能"),43:("良","能勝任所從事的工作","支持其工作適能"),46:("良","能勝任所從事的工作","支持其工作適能"),47:("優","能很好地勝任所從事的工作","維持其工作適能"),49:("優","能很好地勝任所從事的工作","維持其工作適能")}
        for total, values in expected.items(): self.assertEqual((calculate_work_ability_level(total),calculate_work_ability_meaning(total),calculate_work_ability_measure(total)), values)
