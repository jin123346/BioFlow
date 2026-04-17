import pandas as pd
from utils.text_utils import normalize_text, normalize_scientific_name

class SpeciesMatchService:
    def __init__(self,species_master_df: pd.DataFrame):
        self.species_master_df = species_master_df.copy()
        self._prepare_master()
        
    def _prepare_master(self):
        df= self.species_master_df
        
        df["__norm_kor__"] = df["국명"].apply(normalize_text) if "국명 (잘라내서 붙이기 안됨 복사해서 붙이기는 허용)" in df.columns else ""
        df["__norm_sci__"] = df["학명"].apply(normalize_text) if "학명 (국며입력시 자동생성)" in df.columns else ""
        
        self.species_master_df = df
        
        self.kor_map = {}
        self.sci_map = {}
        
        for _, row in df.iterrows():
            species_id = row.get("종학명정보ID","")
            kor_name  = row.get("__norm_kor__","")       
            sci_name  = row.get("__norm_sci__","")
            
            if kor_name:
                self.kor_map.setdefault(kor_name,[]).append(row)
            if sci_name:
                self.sci_map.setdefault(sci_name,[]).append(row)
                
        
    def match_species_id(self, input_df:pd.DataFrame) -> pd.DataFrame:
        result_df = input_df.copy()
        
        result_df["matched_species_id"]=""
        result_df["matched_by"]=""
        result_df["matched_name"]=""
        result_df["matched_sci_name"]=""
        result_df["matched_status"]=""
        
        for idx, row in result_df.iterrows():
            existing_id = normalize_text(row.get("종고유ID",""))
            kor_name = normalize_text(row.get("국명 (잘라내서 붙이기 안됨 복사해서 붙이기는 허용)",""))
            sci_name= normalize_scientific_name(row.get("학명 (국며입력시 자동생성)",""))
            
            #1.기존 종고유 ID가 있으면 유지 
            # result_df.at[idx,"matched_species_id"]= matched.get("종학명정보ID", "")
            # result_df.at[idx,"matched_by"]="기존Id"
            # result_df.at[idx,"matched_name"]=