TARGET_SHEET_NAME = "국명이없거나 학명정보가잘못된것 기재 "
INPUT_HEADER_ROW = 1
OUTPUT_GUIDE_ROW = 1
OUTPUT_HEADER_ROW = 2
OUTPUT_DATA_START_ROW = 3
BASIC_SHEET_NAME="정보입력"

TAXON_CODE_RULES = [
    ("식물계", "PL"),
    ("Plantae", "PL"),
    ("곤충강", "IN"),
    ("Insecta", "IN"),
    ("조강", "BI"),
    ("Aves", "BI"),
    ("조기강", "PI"),
    ("Actinopterygii", "PI"),
    ("포유강", "MM"),
    ("Mammalia", "MM"),
]