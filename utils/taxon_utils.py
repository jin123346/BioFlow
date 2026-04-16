def detect_taxon_code(row: pd.Series) -> str:
    for field in [
        "분류코드",
        "계 국문",
        "강 국문",
        "문 국문",
        "계 영명",
        "강 영명",
        "문 영명",
        "Family",
        "과국명",
    ]:
        value = normalize_text(row.get(field, ""))
        if not value:
            continue
        for keyword, code in TAXON_CODE_RULES:
            if keyword.lower() in value.lower():
                return code
    return ""