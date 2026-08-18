import unittest
from analyses.framingham_v1 import *
class Scores(unittest.TestCase):
 def test_age_boundaries(self):
  male={34:-1,35:0,39:0,40:1,44:1,45:2,49:2,50:3,54:3,55:4,59:4,60:5,64:5,65:6,69:6,70:7,100:7};female={34:-9,35:-4,39:-4,40:0,44:0,45:3,49:3,50:6,54:6,55:7,59:7,60:8,100:8}
  for age,score in male.items():self.assertEqual(calculate_age_score("男",age),score)
  for age,score in female.items():self.assertEqual(calculate_age_score("女",age),score)
 def test_blood_pressure(self):
  for value,score in {129:0,130:1,139:1,140:2,159:2,160:3}.items():self.assertEqual(calculate_blood_pressure_score("男",value,80),score)
  for value,score in {84:0,85:1,89:1,90:2,99:2,100:3}.items():self.assertEqual(calculate_blood_pressure_score("男",110,value),score)
  for value,score in {119:-3,120:0,139:0,140:2,159:2,160:3}.items():self.assertEqual(calculate_blood_pressure_score("女",value,79),score)
  for value,score in {79:-3,80:0,84:0,85:0,89:0,90:2,99:2,100:3}.items():self.assertEqual(calculate_blood_pressure_score("女",110,value),score)
  self.assertEqual(calculate_blood_pressure_score("男",120,95),2)
 def test_cholesterol_and_hdl(self):
  vals=[159,160,199,200,239,240,279,280]
  self.assertEqual([calculate_cholesterol_score("男",x) for x in vals],[-3,0,0,1,1,2,2,3]);self.assertEqual([calculate_cholesterol_score("女",x) for x in vals],[-2,0,0,1,1,1,1,3])
  vals=[34,35,44,45,49,50,59,60]
  self.assertEqual([calculate_hdl_score("男",x) for x in vals],[2,1,1,0,0,0,0,-2]);self.assertEqual([calculate_hdl_score("女",x) for x in vals],[5,2,2,1,1,0,0,-3])
 def test_risks(self):
  male=[-1,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14];expected=["2%","3%","3%","4%","5%","7%","8%","10%","13%","16%","20%","25%","31%","37%","45%","≥53%"]
  self.assertEqual([calculate_total_risk("男",x) for x in male],expected)
  female=list(range(-2,18)); exp=["1%","2%","2%","2%","3%","3%","4%","4%","5%","6%","7%","8%","10%","11%","13%","15%","18%","20%","24%","≥27%"]
  self.assertEqual([calculate_total_risk("女",x) for x in female],exp)
 def test_incidence_levels_comparison(self):
  self.assertEqual(calculate_age_incidence("女",30),"<1%");self.assertEqual(calculate_age_incidence("男",70),"14%")
  self.assertEqual([calculate_risk_level(x) for x in ["9%","10%","20%","21%","30%","31%"]],["低度","中度","中度","高度","高度","極高"])
  pairs=[("2%","3%","較低"),("3%","3%","一樣"),("4%","3%","較高"),("1%","<1%","較高"),("2%","<1%","較高"),("≥53%","≥27%","較高")]
  for a,b,c in pairs:self.assertEqual(compare_with_age_incidence(a,b),c)
