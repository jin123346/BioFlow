# services/species_compare_service.py

import pandas as pd
from utils.text_utils import normalize_text


class SpeciesCompareService:

    def compare(self, master_df: pd.DataFrame, candidate_df: pd.DataFrame) -> pd.DataFrame:
        master_df = master_df.copy()
        candidate_df = candidate_df.copy()

        master_df["학명"] = master_df["학명"].apply(normalize_text)
        master_df["국명"] = master_df["국명"].apply(normalize_text)

        candidate_df["학명"] = candidate_df["학명"].apply(normalize_text)
        candidate_df["국명"] = candidate_df["국명"].apply(normalize_text)

        result_rows = []

        for _, row in candidate_df.iterrows():
            cand_scnm = row.get("학명")
            cand_cnnm = row.get("국명")

            matched_by_scnm = master_df[master_df["학명"] == cand_scnm] if cand_scnm else pd.DataFrame()
            matched_by_cnnm = master_df[master_df["국명"] == cand_cnnm] if cand_cnnm else pd.DataFrame()

            matched_row = None
            status = ""
            match_basis = ""

            if not cand_cnnm and not cand_scnm:
                status = "식별정보부족"
                match_basis = "학명/국명 없음"

            elif len(matched_by_scnm) == 0 and len(matched_by_cnnm) == 0:
                status = "신규"
                match_basis = "미매칭"

            elif len(matched_by_scnm) == 1:
                matched_row = matched_by_scnm.iloc[0]
                status = "업데이트"
                match_basis = "학명 일치"

            elif len(matched_by_cnnm) == 1:
                matched_row = matched_by_cnnm.iloc[0]
                status = "업데이트"
                match_basis = "국명 일치"

            else:
                status = "검토필요"
                match_basis = "중복 또는 부분일치"

            result_row = row.to_dict()
            result_row["status"] = status
            result_row["match_basis"] = match_basis

            if matched_row is not None:
                result_row["matched_specs_essnt_info_id"] = matched_row.get("종학명정보ID", "")
                result_row["matched_scnm"] = matched_row.get("학명", "")
                result_row["matched_cnnm"] = matched_row.get("국명", "")

            result_rows.append(result_row)

        return pd.DataFrame(result_rows)