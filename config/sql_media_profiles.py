"""NARIS mappings with the column types and lengths supplied by the user."""

from copy import deepcopy


def column(name, aliases=(), data_type="VARCHAR2", required=False, value=None, computed=None, length=None):
    return {"name": name, "aliases": list(aliases), "type": data_type,
            "required": required, "constant": value, "computed": computed,
            "length": length, "length_unit": "BYTE", "source": None, "filename_extension": "", "folder_lookup": False}


def image_profile(label, table, key, aliases, specimen=False):
    columns = [
        column(key, aliases, "VARCHAR2", True),
        column("IMAGE_SEQ", data_type="CHAR", required=True, value="1", computed="sequence", length=1),
        column("IMAGE_TPCD", ("이미지구분코드",), required=True, value="03"),
        column("STORE_FILE_NM", ("파일명", "저장파일명"), required=True),
        column("ORGIL_FILE_NM", ("원본파일명",)),
        column("FILE_MGNTD", ("파일크기",), "NUMBER"),
        column("FILE_STORE_CRSE", required=True, computed="nas"),
    ]
    columns.append(column("FILE_DTL_CONT", ("이미지 설명", "이미지설명")))
    return {"label": label, "table": table, "key": key, "sequence": "IMAGE_SEQ",
            "sequence_mode": "fixed", "sequence_value": "1", "sequence_confirmed": True,
            "nas_template": f"/NAS/naris/{'smpl' if specimen else 'obsrv'}/naris/{{기관코드}}/", "columns": columns, "media_names": ["이미지", label, "image"]}


def av_profile(video):
    prefix = "MVPCT" if video else "SOUND"
    name = "동영상" if video else "사운드"
    columns = [
        column("SPECS_ESSNT_ID", ("종고유ID",), "NUMBER", True),
        column(f"{prefix}_INFO_SEQ", data_type="NUMBER", required=True, computed="sequence"),
        column(f"LVBNG_{prefix}_SUBJC_KOR", (f"{name} 제목",)),
        column(f"LVBNG_{prefix}_SUBJC_ENG", ("동영상 제목 영문", "동영상 영문 제목") if video else (), computed=None if video else "null"),
        column(f"LVBNG_{prefix}_CONT_KOR", (f"{name} 상세설명",)),
        column(f"LVBNG_{prefix}_CONT_ENG", ("동영상 상세설명 영문", "동영상 영문 설명") if video else (), computed=None if video else "null"),
        column(f"LVBNG_{prefix}_PLAY_TM", (f"{name} 재생시간",), computed="duration", length=5),
    ]
    if video:
        columns.extend([
            column("LVBNG_MVPCT_USE_YN", ("사용 여부", "사용여부"), required=True, value="Y"),
            column("LVBNG_MVPCT_RPRSN_IMAGE", ("썸네일 파일명", "썸네일파일명", "동영상 대표이미지"), required=True, length=100),
        ])
    columns.extend([
        column("PH500_FILE_NM" if video else "PH192_FILE_NM", (f"{name} 파일명", "파일명"), required=True),
        column("FILE_STORE_CRSE", required=True, computed="nas"),
    ])
    if not video:
        columns.append(column("LVBNG_SOUND_USE_YN", ("사용 여부", "사용여부"), required=True, value="Y"))
    columns.extend([
        column("FRST_REGISTER_ID", required=True, value="narisorg"),
        column("FRST_REGIST_PNTTM", data_type="DATE", required=True, computed="sysdate"),
        column("LAST_UPDUSR_ID", computed="null"), column("LAST_UPDT_PNTTM", data_type="DATE", computed="null"),
    ])
    return {"label": "관찰동영상" if video else "관찰소리",
            "table": f"TB_SBD_LVBNG_{prefix}_INFO", "key": "SPECS_ESSNT_ID",
            "sequence": f"{prefix}_INFO_SEQ", "sequence_mode": "max_per_key", "sequence_confirmed": True,
            "sequence_value": "1", "columns": columns,
            "nas_template": f"/NAS/naris/{'movie' if video else 'sound'}/naris/specs/{{SPECS_ESSNT_ID}}/",
            "media_names": ["동영상", "관찰동영상", "video"] if video else ["사운드", "소리", "관찰소리", "sound", "audio"]}


PROFILES = {
    "observation_image": image_profile("관찰이미지", "TB_SBD_OBSRV_IMAGE_INFO", "OBSRV_INFO_SEQ", ("관찰정보순번",)),
    "observation_video": av_profile(True),
    "observation_sound": av_profile(False),
    "specimen_image": image_profile("표본이미지", "TB_SBD_SPECS_IMAGE_INFO", "SMPL_INFO_SEQ", ("표본정보순번",), True),
}

# Confirmed schema supplied by the user; shared by defaults and saved mappings.
for profile in PROFILES.values():
    for item in profile["columns"]:
        name = item["name"]
        if name in ("OBSRV_INFO_SEQ", "SMPL_INFO_SEQ"):
            item.update(type="CHAR", length=20)
        elif name == "IMAGE_SEQ":
            item.update(type="CHAR", length=1)
        elif name == "LVBNG_SOUND_USE_YN":
            item.update(type="CHAR", length=1)
        elif item["type"] in ("CHAR", "VARCHAR2"):
            length = (2 if name in ("IMAGE_TPCD", "LVBNG_MVPCT_USE_YN") else
                      255 if name == "FILE_STORE_CRSE" else
                      4000 if name == "FILE_DTL_CONT" or "_CONT_" in name else
                      2000 if "_SUBJC_" in name else
                      20 if name.endswith("_PLAY_TM") or name in ("FRST_REGISTER_ID", "LAST_UPDUSR_ID") else 100)
            item.update(length=length)


def default_profile(kind):
    result = deepcopy(PROFILES[kind])
    result.update(kind=kind, schema_confirmed=True, type_column=None, code_column=None,
                  target_codes=[], other_codes=[], all_rows_confirmed=False,
                  use_codes=["Y"] if kind in ("observation_video", "observation_sound") else [],
                  use_codes_confirmed="image" in kind or kind in ("observation_video", "observation_sound"), database_number_column=None,
                  vernacular_name_column=None, scientific_name_column=None,
                  organization_code_column=None, organization_code="",
                  media_lookup_dir="", lookup_recursive=True, lookup_skip_previews=True,
                  encoding="utf-8", lock_table=True)
    return result
