from copy import deepcopy
from datetime import datetime, time, timedelta
from pathlib import Path
import tempfile
import unittest

from config.sql_media_profiles import PROFILES, default_profile
from services.sql_excel_service import SheetData, sql_script
from services.sql_generation_service import auto_map_profile, duration_hhmm, generate_sql, nas_path, number


def record(media="동영상", **extra):
    values = {"종고유ID": "29848", "관찰정보순번": "8123", "표본정보순번": "SMPL-001", "미디어타입": media,
              "파일명": "image.jpg", "원본파일명": "original", "파일크기": "1234",
              "동영상 제목": "찔레꽃 O'Brien", "동영상 상세설명": "설명\n둘째 줄 & 기록",
              "동영상 재생시간": "00:01:25", "동영상 파일명": "video.mp4",
              "썸네일 파일명": "thumbnail.jpg", "데이터베이스번호": "NSMK-OM-BI-260100001",
              "사운드 제목": "소리", "사운드 상세설명": "설명", "사운드 재생시간": "01:25:30",
              "사운드 파일명": "sound.wav", "사용 여부": "Y"}
    values.update(extra)
    return values


def profile_for(kind, headers):
    profile = auto_map_profile(default_profile(kind), headers)
    # Test-only schema values, not assertions about the production DB.
    profile.update(schema_confirmed=True, use_codes=["Y", "N"], use_codes_confirmed=True)
    if "image" in kind:
        profile["nas_template"] = "/NAS/test/images/"
    for column in profile["columns"]:
        if column["type"] in ("VARCHAR2", "CHAR") and column["length"] is None:
            column["length"] = 1000
    return profile


def sheet_of(*rows):
    return SheetData("미디어", 2, list(rows[0]) if rows else [], [(i + 3, row) for i, row in enumerate(rows)])


class SQLGenerationTests(unittest.TestCase):
    def test_confirmed_schema_and_image_key_length(self):
        for kind in PROFILES:
            profile = default_profile(kind)
            for col in profile["columns"]:
                if col["type"] in ("CHAR", "VARCHAR2"):
                    self.assertGreater(col["length"], 0, (kind, col["name"]))
            if "image" not in kind:
                continue
            header = "관찰정보순번" if kind == "observation_image" else "표본정보순번"
            row = record("이미지", **{header: "OBSR_000000000764001", "기관코드": "00123"})
            profile = auto_map_profile(profile, row)
            key = profile["columns"][0]
            self.assertEqual((key["type"], key["length"]), ("CHAR", 20))
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "SUCCESS", result.error)
            self.assertIn("'OBSR_000000000764001'", result.sql)
            row[header] = "X" * 21
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "ERROR")
            self.assertIn("21/20", result.error)

    def test_image_keys_are_fixed_varchar_and_preserve_text_and_leading_zeroes(self):
        for kind, header, key in (("observation_image", "관찰정보순번", "OBSRV_INFO_SEQ"),
                                  ("specimen_image", "표본정보순번", "SMPL_INFO_SEQ")):
            for value in ("OBS-A001", "000123"):
                row = record("이미지", **{header: value, "기관코드": "00123"})
                profile = profile_for(kind, row)
                profile["nas_template"] = ""
                profile["columns"][0]["type"] = "NUMBER"
                result = generate_sql(sheet_of(row), profile)[0]
                self.assertEqual(result.status, "SUCCESS", result.error)
                self.assertEqual(result.key, value)
                self.assertIn(key, result.sql)
                self.assertIn(f"'{value}'", result.sql)
                self.assertIn("/naris/00123/", result.sql)

    def test_image_nas_paths_use_mapped_or_fixed_organization_code(self):
        for kind, folder in (("observation_image", "obsrv"), ("specimen_image", "smpl")):
            with self.subTest(kind=kind):
                row = record("이미지", 기관코드="00123")
                profile = profile_for(kind, row)
                profile["nas_template"] = PROFILES[kind]["nas_template"]
                profile["organization_code"] = "OTHER"
                result = generate_sql(sheet_of(row), profile)[0]
                self.assertEqual(result.status, "SUCCESS", result.error)
                self.assertIn(f"'/NAS/naris/{folder}/naris/00123/'", result.sql)
                profile["organization_code_column"] = None
                result = generate_sql(sheet_of(row), profile)[0]
                self.assertIn(f"'/NAS/naris/{folder}/naris/00123/'", result.sql)
                for invalid in ("", "../bad", "a/b", "a\\b", ".."):
                    row["기관코드"] = invalid
                    result = generate_sql(sheet_of(row), profile)[0]
                    self.assertEqual(result.status, "ERROR")
                    self.assertIn("기관코드", result.error)

    def test_invalid_image_key_reports_mapping_without_false_missing_nas(self):
        row = record("이미지", 관찰정보순번="관찰정보순번", 파일크기="파일크기")
        profile = profile_for("observation_image", row)
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "ERROR")
        self.assertNotIn("OBSRV_INFO_SEQ:", result.error)
        self.assertIn("매핑: 파일크기", result.error)
        self.assertNotIn("FILE_STORE_CRSE: 필수값 누락", result.error)
        self.assertNotIn("IMAGE_SEQ:", result.error)

    def test_image_sequence_one_overrides_stale_settings_and_excel_mapping(self):
        for kind in ("observation_image", "specimen_image"):
            for mode in ("source", "max_per_key", "fixed"):
                with self.subTest(kind=kind, mode=mode):
                    row = record("이미지", IMAGE_SEQ="not a number")
                    profile = profile_for(kind, row)
                    profile.update(sequence_mode=mode, sequence_value="99", sequence_confirmed=False)
                    seq = next(col for col in profile["columns"] if col["name"] == "IMAGE_SEQ")
                    seq.update(source="IMAGE_SEQ", constant="99", type="NUMBER", length=100)
                    result = generate_sql(sheet_of(row), profile)[0]
                    self.assertEqual(result.status, "SUCCESS", result.error)
                    self.assertEqual(result.sequence, "1")
                    values = result.sql.split("VALUES (", 1)[1].split(", ")
                    self.assertEqual(values[1].strip(), "'1'")
                    self.assertEqual(profile["sequence_value"], "99")
                    mapped = auto_map_profile(profile, row)
                    self.assertEqual(mapped["sequence_mode"], "fixed")
                    self.assertEqual(mapped["sequence_value"], "1")
                    self.assertIsNone(seq["source"])
                    self.assertEqual(seq["type"], "CHAR")
                    self.assertEqual(seq["length"], 1)

    def test_all_four_table_profiles_and_image_fixed_rules(self):
        for kind in PROFILES:
            row = record("이미지" if "image" in kind else "동영상" if "video" in kind else "사운드")
            profile = profile_for(kind, row)
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "SUCCESS", result.error)
            self.assertIn("INSERT INTO " + PROFILES[kind]["table"], result.sql)
            self.assertTrue(result.sql.endswith(";"))
            self.assertNotIn("\n", result.sql)
            self.assertNotIn("\r", result.sql)
            if "image" in kind:
                self.assertEqual(result.sequence, "1")
                self.assertIn("'03'", result.sql)
                self.assertNotIn("MAX(", result.sql)
            else:
                self.assertIn(f"NVL(MAX({profile['sequence']}), 0) + 1", result.sql)
                self.assertIn("WHERE SPECS_ESSNT_ID = 29848", result.sql)
                self.assertIn("'narisorg'", result.sql)
                self.assertIn("SYSDATE", result.sql)

    def test_video_thumbnail_description_escaping_duration_and_nas(self):
        row = record()
        result = generate_sql(sheet_of(row), profile_for("observation_video", row))[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("O''Brien", result.sql)
        self.assertIn("CHR(10)", result.sql)
        self.assertIn("<br> : NSMK-OM-BI-260100001", result.sql)
        self.assertIn("'thumbnail.jpg'", result.sql)
        self.assertIn("'00:01'", result.sql)
        self.assertIn("'/NAS/naris/movie/naris/specs/29848/'", result.sql)
        self.assertIn("SET DEFINE OFF", sql_script([result], profile_for("observation_video", row)))

    def test_video_use_yn_is_confirmed_fixed_y_even_with_stale_mapping(self):
        row = record()
        del row["사용 여부"]
        profile = default_profile("observation_video")
        profile.update(use_codes=[], use_codes_confirmed=False)
        use_column = next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_USE_YN")
        use_column.update(source="없는 컬럼", constant="N", required=False)

        result = generate_sql(sheet_of(row), auto_map_profile(profile, row))[0]

        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("LVBNG_MVPCT_USE_YN", result.sql)
        self.assertIn("'00:01', 'Y', 'thumbnail.jpg'", result.sql)

    def test_sound_use_yn_is_confirmed_fixed_y_even_with_stale_mapping(self):
        row = record("사운드")
        del row["사용 여부"]
        profile = default_profile("observation_sound")
        profile.update(use_codes=[], use_codes_confirmed=False)
        use_column = next(col for col in profile["columns"] if col["name"] == "LVBNG_SOUND_USE_YN")
        use_column.update(source="없는 컬럼", constant="N", required=False)

        result = generate_sql(sheet_of(row), auto_map_profile(profile, row))[0]

        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("LVBNG_SOUND_USE_YN", result.sql)
        self.assertIn("'sound.wav', '/NAS/naris/sound/naris/specs/29848/', 'Y'", result.sql)

    def test_observation_image_never_substitutes_species_id_for_observation_key(self):
        row = record("이미지")
        del row["관찰정보순번"]
        result = generate_sql(sheet_of(row), profile_for("observation_image", row))[0]
        self.assertEqual(result.status, "ERROR")
        self.assertIn("OBSRV_INFO_SEQ", result.error)
        self.assertEqual(result.sql, "")

    def test_missing_thumbnail_is_error_and_not_invented_from_video_name(self):
        row = record()
        del row["썸네일 파일명"]
        profile = profile_for("observation_video", row)
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "ERROR")
        self.assertIn("LVBNG_MVPCT_RPRSN_IMAGE", result.error)
        next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_RPRSN_IMAGE")["constant"] = "manual.jpg"
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("'manual.jpg'", result.sql)

    def test_fixed_image_sequence_duplicates_mark_all_conflicting_rows(self):
        first = record("이미지")
        second = record("이미지", 파일명="second.jpg")
        results = generate_sql(sheet_of(first, second), profile_for("observation_image", first))
        self.assertEqual([result.status for result in results], ["ERROR", "ERROR"])
        self.assertTrue(all("SEQ 중복" in result.error for result in results))

    def test_video_same_species_multiple_files_allow_sequential_max(self):
        first = record()
        second = record(**{"동영상 파일명": "second.mp4"})
        results = generate_sql(sheet_of(first, second), profile_for("observation_video", first))
        self.assertEqual([result.status for result in results], ["SUCCESS", "SUCCESS"])
        script = sql_script(results, profile_for("observation_video", first))
        self.assertLess(script.index("'video.mp4'"), script.index("'second.mp4'"))
        self.assertIn("LOCK TABLE TB_SBD_LVBNG_MVPCT_INFO IN EXCLUSIVE MODE;", script)
        duplicate = generate_sql(sheet_of(first, first), profile_for("observation_video", first))
        self.assertTrue(all(result.status == "ERROR" for result in duplicate))

    def test_filter_skips_other_types_and_rejects_unconfirmed_ob_code(self):
        video, photo = record(), record("이미지")
        profile = profile_for("observation_video", video)
        result = generate_sql(sheet_of(video, photo), profile)
        self.assertEqual([item.status for item in result], ["SUCCESS", "SKIP"])
        video["미디어코드"] = "OB"
        profile = profile_for("observation_video", video)
        result = generate_sql(sheet_of(video), profile)[0]
        self.assertEqual(result.status, "ERROR")
        self.assertIn("미확정", result.error)
        profile["target_codes"] = ["OB"]
        self.assertEqual(generate_sql(sheet_of(video), profile)[0].status, "SUCCESS")

    def test_no_filter_uses_selected_sql_type_and_can_be_disabled(self):
        row = record()
        del row["미디어타입"]
        profile = profile_for("observation_video", row)
        self.assertTrue(profile["all_rows_confirmed"])
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "SUCCESS")
        profile["all_rows_confirmed"] = False
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")

    def test_schema_unknown_lengths_codes_and_invalid_values_are_errors(self):
        row = record()
        profile = auto_map_profile(default_profile("observation_video"), row)
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertNotIn("최대 길이가 미확정", result.error)
        for updates in ({"종고유ID": "bad"}, {"썸네일 파일명": "x" * 101},
                        {"동영상 재생시간": "00:99"},
                        {"동영상 파일명": "C:\\video.mp4"}):
            invalid = record(**updates)
            result = generate_sql(sheet_of(invalid), profile_for("observation_video", invalid))[0]
            self.assertEqual(result.status, "ERROR", updates)
            self.assertIn("엑셀 3행", result.error)

    def test_duration_string_excel_time_and_fraction(self):
        for value, expected in (("00:00:45", "00:00"), ("00:01:25", "00:01"),
                                ("01:25:30", "01:25"), ("02:10", "02:10"), (None, None),
                                (time(1, 25, 59), "01:25"), (timedelta(hours=25, minutes=3), "25:03"),
                                (datetime(2026, 1, 1, 2, 3), "02:03"), (0.5, "12:00")):
            self.assertEqual(duration_hhmm(value), expected)
        for value in ("25:99", "00:00:60", "bad", -1, "100:00"):
            with self.assertRaises(ValueError):
                duration_hhmm(value)

    def test_nas_paths_are_not_local_paths_or_arbitrary_placeholder_expressions(self):
        self.assertEqual(nas_path("/NAS//movie/{SPECS_ESSNT_ID}", {"SPECS_ESSNT_ID": "12"}), "/NAS/movie/12/")
        for value in ("", "C:\\NAS", "/NAS/../other", "/NAS/{unknown}/", "/NAS/{SPECS_ESSNT_ID.__class__}/"):
            with self.assertRaises(ValueError):
                nas_path(value, {"SPECS_ESSNT_ID": "12"})

    def test_mandatory_columns_and_thumbnail_limit_cannot_be_relaxed(self):
        row = record()
        profile = profile_for("observation_video", row)
        next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_RPRSN_IMAGE")["length"] = 101
        invalid = record(**{"썸네일 파일명": "x" * 101})
        result = generate_sql(sheet_of(invalid), profile)[0]
        self.assertEqual(result.status, "ERROR")
        self.assertIn("100", result.error)

    def test_numeric_precision_and_sql_injection_are_validated(self):
        self.assertEqual(number("9" * 38), "9" * 38)
        for value in ("9" * 39, "1; DROP TABLE T", float("inf"), "1e126", "1e-131"):
            with self.assertRaises(ValueError):
                number(value)
        self.assertEqual(number("29848.0", True), "29848")

    def test_multibyte_db_lengths_and_required_null_are_validated(self):
        row = record(**{"동영상 제목": "가" * 667})
        profile = profile_for("observation_video", row)
        title = next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_SUBJC_KOR")
        title.update(length=3, length_unit="BYTE")
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")
        row["동영상 제목"] = "가" * 666
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "SUCCESS")
        next(col for col in profile["columns"] if col["computed"] == "null")["required"] = True
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")

    def test_empty_video_titles_and_descriptions_fall_back_by_language(self):
        row = record(**{"동영상 제목": "", "동영상 상세설명": None, "국명": "찔레꽃", "학명": "Rosa multiflora"})
        profile = profile_for("observation_video", row)
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("'찔레꽃'", result.sql)
        self.assertIn("'찔레꽃<br> : NSMK-OM-BI-260100001'", result.sql)
        self.assertEqual(result.sql.count("'Rosa multiflora'"), 2)

    def test_existing_video_fields_are_kept_and_only_blank_fields_fall_back(self):
        row = record(**{"국명": "찔레꽃", "학명": "Rosa multiflora", "동영상 제목 영문": "Existing title",
                        "동영상 상세설명 영문": "Existing description"})
        result = generate_sql(sheet_of(row), profile_for("observation_video", row))[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("O''Brien", result.sql)
        self.assertIn("CHR(10)", result.sql)
        self.assertIn("'Existing title'", result.sql)
        self.assertIn("'Existing description'", result.sql)
        self.assertNotIn("'Rosa multiflora'", result.sql)

    def test_video_name_fallback_supports_saved_null_mapping_and_custom_name_columns(self):
        row = record(**{"동영상 제목": "", "동영상 상세설명": "", "이름": "찔레꽃", "학술명": "Rosa multiflora"})
        profile = profile_for("observation_video", row)
        for column in profile["columns"]:
            if column["name"] in ("LVBNG_MVPCT_SUBJC_ENG", "LVBNG_MVPCT_CONT_ENG"):
                column["computed"] = "null"
        profile = auto_map_profile(profile, row)
        profile.update(vernacular_name_column="이름", scientific_name_column="학술명")
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertIn("'찔레꽃'", result.sql)
        self.assertEqual(result.sql.count("'Rosa multiflora'"), 2)

    def test_sound_english_null_behavior_is_unchanged(self):
        row = record("사운드", **{"국명": "찔레꽃", "학명": "Rosa multiflora"})
        result = generate_sql(sheet_of(row), profile_for("observation_sound", row))[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertNotIn("Rosa multiflora", result.sql)

    def test_database_number_filename_mapping_accepts_per_column_extensions(self):
        for kind, field, extension in (
            ("observation_image", "STORE_FILE_NM", ".jpg"),
            ("specimen_image", "STORE_FILE_NM", ".png"),
            ("observation_video", "PH500_FILE_NM", ".mp4"),
            ("observation_sound", "PH192_FILE_NM", ".wav"),
        ):
            row = record("이미지" if "image" in kind else "동영상" if "video" in kind else "사운드")
            profile = profile_for(kind, row)
            column = next(col for col in profile["columns"] if col["name"] == field)
            column.update(source="데이터베이스번호", filename_extension=extension)
            if kind == "observation_video":
                thumb = next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_RPRSN_IMAGE")
                thumb.update(source="데이터베이스번호", filename_extension="jpg")
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "SUCCESS", result.error)
            self.assertEqual(result.filename, row["데이터베이스번호"] + extension)
            self.assertIn("'" + result.filename + "'", result.sql)
            if kind == "observation_video":
                self.assertIn("'" + row["데이터베이스번호"] + ".jpg'", result.sql)

    def test_extensions_are_optional_and_existing_extensions_are_kept(self):
        row = record("이미지")
        row["파일명"] = "already.PNG"
        profile = profile_for("specimen_image", row)
        column = next(col for col in profile["columns"] if col["name"] == "STORE_FILE_NM")
        column["filename_extension"] = ".jpg"
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "SUCCESS", result.error)
        self.assertEqual(result.filename, "already.PNG")
        column.update(source="데이터베이스번호", filename_extension="")
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.filename, row["데이터베이스번호"])
        column.update(source=None, constant="fixed", filename_extension=".jpg")
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].filename, "fixed.jpg")

    def test_invalid_extensions_empty_filenames_and_final_length_are_checked(self):
        row = record()
        profile = profile_for("observation_video", row)
        thumb = next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_RPRSN_IMAGE")
        thumb.update(filename_extension="../jpg")
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")
        thumb["filename_extension"] = ".jpg"
        row["썸네일 파일명"] = ""
        self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")
        row["썸네일 파일명"] = "x" * 98
        result = generate_sql(sheet_of(row), profile)[0]
        self.assertEqual(result.status, "ERROR")
        self.assertIn("길이 초과", result.error)

    def test_folder_matches_actual_names_for_all_media_and_video_thumbnail(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            subfolder = root / "29848"
            subfolder.mkdir()
            db = record()["데이터베이스번호"]
            for extension in (".JPG", ".mp4", ".wav"):
                (subfolder / (db + extension)).write_bytes(b"unchanged")
            before = {str(path): path.read_bytes() for path in subfolder.iterdir()}
            for kind, field, extension in (
                ("observation_image", "STORE_FILE_NM", ".JPG"),
                ("specimen_image", "STORE_FILE_NM", ".JPG"),
                ("observation_video", "PH500_FILE_NM", ".mp4"),
                ("observation_sound", "PH192_FILE_NM", ".wav"),
            ):
                row = record("이미지" if "image" in kind else "동영상" if "video" in kind else "사운드")
                profile = profile_for(kind, row)
                profile["media_lookup_dir"] = folder
                next(col for col in profile["columns"] if col["name"] == field)["folder_lookup"] = True
                if kind == "observation_video":
                    thumb = next(col for col in profile["columns"] if col["name"] == "LVBNG_MVPCT_RPRSN_IMAGE")
                    thumb.update(source=None, constant=None, folder_lookup=True)
                result = generate_sql(sheet_of(row), profile)[0]
                self.assertEqual(result.status, "SUCCESS", result.error)
                self.assertEqual(result.filename, db + extension)
                if kind == "observation_video":
                    self.assertIn("'" + db + ".JPG'", result.sql)
            self.assertEqual({str(path): path.read_bytes() for path in subfolder.iterdir()}, before)

    def test_folder_ambiguous_images_require_extension_filter(self):
        with tempfile.TemporaryDirectory() as folder:
            row = record("이미지")
            db = row["데이터베이스번호"]
            for extension in (".jpg", ".png"):
                (Path(folder) / (db + extension)).write_bytes(b"image")
            profile = profile_for("specimen_image", row)
            profile["media_lookup_dir"] = folder
            column = next(col for col in profile["columns"] if col["name"] == "STORE_FILE_NM")
            column["folder_lookup"] = True
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "ERROR")
            self.assertIn("후보가 2개", result.error)
            column["filename_extension"] = ".jpg"
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "SUCCESS", result.error)
            self.assertEqual(result.filename, db + ".jpg")

    def test_folder_lookup_excludes_preview_subfolder_and_can_disable_recursion(self):
        with tempfile.TemporaryDirectory() as folder:
            row = record()
            db = row["데이터베이스번호"]
            subfolder = Path(folder) / "29848"
            preview = subfolder / "썸네일"
            preview.mkdir(parents=True)
            for path in (subfolder / (db + ".mp4"), subfolder / (db + ".jpg"), preview / (db + ".jpg")):
                path.write_bytes(b"media")
            profile = profile_for("observation_video", row)
            profile["media_lookup_dir"] = folder
            for col in profile["columns"]:
                if col["name"] in ("PH500_FILE_NM", "LVBNG_MVPCT_RPRSN_IMAGE"):
                    col["folder_lookup"] = True
            self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "SUCCESS")
            profile["lookup_skip_previews"] = False
            self.assertIn("후보가 2개", generate_sql(sheet_of(row), profile)[0].error)
            profile.update(lookup_skip_previews=True, lookup_recursive=False)
            self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")

    def test_folder_lookup_is_exact_and_uses_explicit_db_column(self):
        with tempfile.TemporaryDirectory() as folder:
            (Path(folder) / "db-10.jpg").write_bytes(b"image")
            row = record("이미지", **{"내 DB 번호": "DB-1"})
            profile = profile_for("specimen_image", row)
            profile.update(media_lookup_dir=folder, database_number_column="내 DB 번호")
            column = next(col for col in profile["columns"] if col["name"] == "STORE_FILE_NM")
            column["folder_lookup"] = True
            self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "ERROR")
            row["내 DB 번호"] = "DB-10"
            result = generate_sql(sheet_of(row), profile)[0]
            self.assertEqual(result.status, "SUCCESS", result.error)
            self.assertEqual(result.filename, "db-10.jpg")
            row["내 DB 번호"] = ""
            self.assertIn("데이터베이스번호", generate_sql(sheet_of(row), profile)[0].error)

    def test_duplicate_same_extension_in_separate_folders_is_not_arbitrarily_selected(self):
        with tempfile.TemporaryDirectory() as folder:
            row = record("이미지")
            for name in ("first", "second"):
                sub = Path(folder) / name
                sub.mkdir()
                (sub / (row["데이터베이스번호"] + ".jpg")).write_bytes(b"image")
            profile = profile_for("specimen_image", row)
            profile["media_lookup_dir"] = folder
            column = next(col for col in profile["columns"] if col["name"] == "STORE_FILE_NM")
            column.update(folder_lookup=True, filename_extension=".jpg")
            self.assertIn("후보가 2개", generate_sql(sheet_of(row), profile)[0].error)
            profile["media_lookup_dir"] = str(Path(folder) / "first")
            self.assertEqual(generate_sql(sheet_of(row), profile)[0].status, "SUCCESS")


if __name__ == "__main__":
    unittest.main()
