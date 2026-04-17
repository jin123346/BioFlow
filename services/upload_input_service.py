import pandas as pd


class UploadInputService:
    def get_sheet_names(self,file_path:str)->list[str]:
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    
    def load_sheet(self,file_path:str, sheet_name: str)->pd.DataFrame:
        df=pd.read_excel(file_path,sheet_name=sheet_name, header=0, dtype=object)
        df.columns = [str(col).strip().replace("\n", " ") for col in df.columns]
        df = df.dropna(how="all") # 전체 행이 비어있는 경우 제거
        return df
    
    