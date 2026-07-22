import tempfile
import unittest
from pathlib import Path

import pandas as pd

from src.rank_jobs import (
    add_component_scores,
    apply_manual_component_fallback,
    extract_manual_company,
    extract_manual_location,
    extract_manual_title,
    get_hard_skip_reason,
    get_location_tier,
    get_manual_final_action_from_ai_score,
    material_files_for_category,
    parse_job_file,
    recommend_application_materials,
    split_manual_job_blocks,
)


def parse_text(text, name="sample.txt"):
    with tempfile.TemporaryDirectory() as temp_dir:
        path = Path(temp_dir) / name
        path.write_text(text, encoding="utf-8")
        return parse_job_file(path)


class RawJobPipelineRegressionTests(unittest.TestCase):
    def score_one(self, row):
        df = pd.DataFrame([row]).fillna("")
        df["hard_skip_reason"] = df.apply(get_hard_skip_reason, axis=1)
        return add_component_scores(df).iloc[0]

    def test_dhl_operations_data_analyst_parses_and_scores_apply(self):
        row = parse_text(
            """
            https://www.linkedin.com/jobs/view/123
            DHL Supply Chain
            Operations Data Analyst
            Milton Keynes, England, United Kingdom
            Full-time
            Save Operations Data Analyst at DHL Supply Chain
            About the job
            SQL, Power BI, Excel, reporting, dashboards, data cleansing, data validation,
            UAT, quality assurance, root cause analysis, stakeholder management, operations.
            """
        )
        scored = self.score_one(row)
        self.assertEqual(scored["job_title"], "Operations Data Analyst")
        self.assertEqual(scored["company"], "DHL Supply Chain")
        self.assertEqual(scored["location_tier"], "Tier 1 - Milton Keynes")
        self.assertEqual(get_hard_skip_reason(scored), "")
        self.assertEqual(scored["recommendation"], "APPLY")
        self.assertGreaterEqual(scored["transferable_experience_score"], 20)

    def test_mk_council_officer_title_is_relevant_not_filtered(self):
        row = parse_text(
            """
            Local Government Chronicle
            Business Intelligence Officer
            Milton Keynes, England, United Kingdom
            Save Business Intelligence Officer at Local Government Chronicle
            We're recruiting a Business Intelligence Officer to join our Children's Services team at Milton Keynes City Council.
            Complex datasets, trend analysis, reporting, statutory data returns, data quality,
            quality assurance, research, evidence and stakeholder communication.
            """
        )
        scored = self.score_one(row)
        self.assertEqual(scored["job_title"], "Business Intelligence Officer")
        self.assertEqual(scored["company"], "Milton Keynes City Council")
        self.assertEqual(scored["role_family"], "Data and BI")
        self.assertEqual(scored["location_tier"], "Tier 1 - Milton Keynes")
        self.assertEqual(get_hard_skip_reason(scored), "")
        self.assertIn(scored["recommendation"], ["APPLY", "HIGH-PRIORITY MANUAL REVIEW", "MANUAL REVIEW"])

    def test_mercedes_hr_business_analyst_graduate_not_rejected_for_hr_or_graduate(self):
        row = parse_text(
            """
            Mercedes AMG High Performance Powertrains
            HR Business Analyst Graduate
            Brixworth
            Save HR Business Analyst Graduate at Mercedes AMG High Performance Powertrains
            Graduate entry-level role supporting data analysis, reporting, HRIS support,
            system testing, process improvement, data policy, compliance and Microsoft Office.
            Degree classification to be confirmed.
            """
        )
        scored = self.score_one(row)
        self.assertEqual(scored["job_title"], "HR Business Analyst Graduate")
        self.assertEqual(scored["location_tier"], "Tier 2 - very high priority commute")
        self.assertEqual(get_hard_skip_reason(scored), "")
        self.assertIn(scored["recommendation"], ["APPLY", "HIGH-PRIORITY MANUAL REVIEW", "MANUAL REVIEW"])
        self.assertGreater(scored["desirable_gap_penalty"], 0)

    def test_vwfs_insight_analyst_contract_is_review_not_hard_skip(self):
        row = parse_text(
            """
            Volkswagen Financial Services (UK)
            Insight Analyst
            Milton Keynes, England, United Kingdom
            Contract
            Save Insight Analyst at Volkswagen Financial Services (UK)
            We're looking for an Insight Executive, known internally as a Commission Insight Executive,
            to analyse large datasets, advanced Excel, reporting, reconciliations, quality checks,
            complaint and case data, Financial Ombudsman Service requests and litigation matters.
            Salary: Starting from £29,854.40 pa
            """
        )
        scored = self.score_one(row)
        self.assertEqual(scored["company"], "Volkswagen Financial Services (UK)")
        self.assertEqual(scored["role_family"], "Data and BI")
        self.assertEqual(scored["location_tier"], "Tier 1 - Milton Keynes")
        self.assertEqual(get_hard_skip_reason(scored), "")
        self.assertIn(scored["recommendation"], ["APPLY", "HIGH-PRIORITY MANUAL REVIEW", "MANUAL REVIEW"])

    def test_santander_business_analyst_finance_control_gap_is_manual_or_stretch(self):
        row = parse_text(
            """
            Santander UK
            Business Analyst | S2 | T&O | Milton Keynes
            Milton Keynes, England, United Kingdom
            Save Business Analyst | S2 | T&O | Milton Keynes at Santander UK
            Financial reporting, dashboards, advanced Excel, reconciliations, budgeting,
            forecasting, variance analysis, accruals, cost tracking, financial controls and process improvement.
            """
        )
        scored = self.score_one(row)
        self.assertEqual(scored["location_tier"], "Tier 1 - Milton Keynes")
        self.assertEqual(get_hard_skip_reason(scored), "")
        self.assertEqual(scored["recommendation"], "MANUAL REVIEW")
        self.assertIn("finance-control gaps", scored["recommendation_reason"])

    def test_parser_keeps_missing_salary_date_url_and_location_as_nonfatal(self):
        row = parse_text(
            """
            Source: Manual
            Job Title: Reporting Executive
            Company: Example Analytics
            Full JD:
            Reporting, Excel, dashboards and stakeholder updates.
            """
        )
        self.assertEqual(row["job_title"], "Reporting Executive")
        self.assertEqual(row["salary"], "")
        self.assertEqual(row["application_deadline"], "")
        self.assertEqual(row["location"], "Unknown")
        self.assertIn("apply link not found", row["manual_parse_notes"])

    def test_multiple_jobs_in_one_file_are_detected(self):
        text = """
        Save Data Analyst at Example One
        Milton Keynes
        SQL, Power BI, reporting.

        Save Finance Analyst at Example Two
        Bedford
        Excel, reconciliations, financial reporting.
        """
        blocks = split_manual_job_blocks(text)
        self.assertEqual(len(blocks), 2)

    def test_london_three_days_is_lower_location_score_than_local(self):
        local = self.score_one(parse_text("Save BI Analyst at Local Co\nMilton Keynes\nPower BI reporting dashboards."))
        london = self.score_one(parse_text("Save BI Analyst at London Co\nLondon\nHybrid three days a week in office. Power BI reporting dashboards."))
        self.assertGreater(local["location_score"], london["location_score"])

    def test_ai_failure_or_missing_ai_can_use_deterministic_fallback(self):
        row = self.score_one(parse_text("Save Data Analyst at Example Co\nMilton Keynes\nSQL Power BI Excel dashboards reporting data validation."))
        df = pd.DataFrame([row]).fillna("")
        df["final_action"] = "Pending AI Review"
        df["final_action_reason"] = "manual full JD awaiting AI review"
        df["hard_skip_reason"] = ""
        updated = apply_manual_component_fallback(df).iloc[0]
        self.assertIn(updated["final_action"], ["Apply If Time", "Manual Review"])
        self.assertTrue(bool(updated["deterministic_fallback_used"]))

    def test_valid_ai_skip_low_score_is_not_rescued_by_deterministic_fallback(self):
        row = self.score_one(parse_text("Save Data Analyst at Example Co\nMilton Keynes\nSQL Power BI Excel dashboards reporting data validation."))
        df = pd.DataFrame([row]).fillna("")
        df["final_action"] = "Skip"
        df["final_action_reason"] = "manual AI score 25; threshold action Skip"
        df["hard_skip_reason"] = ""
        df["ai_final_action"] = "Skip"
        df["ai_fit_score"] = 25
        df["ai_review_source"] = "openai_new"
        df["ai_red_flags"] = ""
        updated = apply_manual_component_fallback(df).iloc[0]
        self.assertEqual(updated["final_action"], "Skip")
        self.assertFalse(bool(updated["deterministic_fallback_used"]))

    def test_local_relevant_manual_job_can_be_apply_if_time_from_score_50(self):
        row = self.score_one(parse_text("Save Reporting Analyst at Local Co\nMilton Keynes\nExcel reporting dashboards stakeholder updates."))
        row["ai_fit_score"] = 50
        self.assertEqual(get_manual_final_action_from_ai_score(row), "Apply If Time")

    def test_application_materials_use_current_cv_filenames_and_specific_category(self):
        cv, cover_letter = material_files_for_category("business_operations_analyst")
        self.assertEqual(cv, "cv_business_operations_analyst.docx")
        self.assertEqual(cover_letter, "cover_letter_business_operations_analyst_ats.docx")
        row = parse_text("Save Junior Business Analyst at Example Co\nMilton Keynes\nRequirements, UAT, process improvement and reporting.")
        materials = recommend_application_materials(row)
        self.assertEqual(materials["cv_category"], "business_operations_analyst")
        self.assertEqual(materials["recommended_cv"], "cv_business_operations_analyst.docx")


if __name__ == "__main__":
    unittest.main()
