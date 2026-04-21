class DBNumberService:
    MEDIA_CODE_MAP = {
        "관찰정보": "OB",
        "이미지": "OI",
        "동영상": "OM",
        "사운드": "OS",
        "표본": "SP",
    }

    def __init__(self, db_number_dao):
        self.db_number_dao = db_number_dao

    def get_media_code(self, sheet_name: str, is_specimen: bool = False) -> str:
        if is_specimen:
            return "SP"
        return self.MEDIA_CODE_MAP.get(sheet_name, "")
    
    def get_taxon_code(self,matched_row: dict) -> str:
        kor_name = str(matched_row.get(""))