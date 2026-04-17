# services/species_compare_service.py

import pandas as pd
from utils.text_utils import normalize_text
from config.constants import TARGET_SHEET_NAME
from PySide6.QtWidgets import QFileDialog



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
    
    def extract_update_candidates(self):
        if not self.update_files:
            self.append_log("업데이트할 파일을 먼저 선택해 주세요.")
            return
        try:
            df = self.candidate_service.extract_candidates(self.update_files, TARGET_SHEET_NAME)
            self.df_candidates = df
            self.append_log(f"업데이트 후보 통합 완료 : 총 {len(df)} 행")
            self.show_dataframe_preview(df)
        except Exception as e:
            self.append_log(f"업데이트 후보 추출 실패: {str(e)}")         
    
    def save_compare_result_to_excel(self):
        if self.df_compare_result is None or self.df_compare_result.empty:
            self.append_log("저장할 비교 결과가 없습니다.")
            return

        file_path, _ = QFileDialog.getSaveFileName(...)
        if not file_path:
            self.append_log("비교 결과 저장이 취소되었습니다.")
            return

        try:
            self.export_service.save_compare_result(self.df_compare_result, file_path)
            self.append_log(f"비교 결과 저장 완료: {file_path}")
        except Exception as e:
            self.append_log(f"비교 결과 저장 실패: {str(e)}")
            
    def build_summary_log(self, result_df: pd.DataFrame) -> str:
        if result_df is None or result_df.empty:
            return "비교 결과가 없습니다."

        return (
            f"비교 완료 - 신규: {(result_df['status'] == '신규').sum()}건, "
            f"업데이트: {(result_df['status'] == '업데이트').sum()}건, "
            f"업데이트(중복): {(result_df['status'] == '업데이트(중복)').sum()}건, "
            f"검토필요: {(result_df['status'] == '검토필요').sum()}건, "
            f"식별정보부족: {(result_df['status'] == '식별정보부족').sum()}건"
        )