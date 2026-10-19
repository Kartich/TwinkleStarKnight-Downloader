import csv
import importlib.util
import tempfile
import unittest
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parents[1]
CORE_PATH = ROOT / "core" / "smart_bundle_downloader_export_v7.py"
SPEC = importlib.util.spec_from_file_location("smart_bundle_test_core", CORE_PATH)
CORE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(CORE)


class CharacterExportPathTests(unittest.TestCase):
    def route(self, path, asset_id="", bundle="", asset_type="", layout="character", variant="", asset_root="", category=""):
        return CORE._route_export_path(
            PurePosixPath(path), layout, asset_id, variant, bundle, asset_type,
            asset_root, category,
        )

    def test_character_leaf_drops_unity_wrapper_chain(self):
        actual = self.route(
            "Chara_Thumb_004_a/AssetBundles/Sprites/Chara/Thumb_004.png",
            asset_id="1118001", asset_type="Sprite"
        )
        self.assertEqual(actual.as_posix(), "111800/1118001/Thumb_004.png")

    def test_explicit_ch_costume_folder_is_retained_without_type_layer(self):
        actual = self.route(
            "ch_1077004/AssetBundles/Sprites/Chara/Body.png",
            asset_id="1077004", asset_type="Sprite"
        )
        self.assertEqual(actual.as_posix(), "107700/1077004/ch_1077004/Body.png")

    def test_explicit_costume_folder_is_retained_without_type_layer(self):
        actual = self.route(
            "1077004/AssetBundles/Spine/1077004.skel",
            asset_id="1077004", asset_type="TextAsset"
        )
        self.assertEqual(actual.as_posix(), "107700/1077004/1077004/1077004.skel")

    def test_un_numbered_resources_are_grouped_only_by_type(self):
        sprite = self.route(
            "Chara_Thumb_004_a/AssetBundles/Sprites/Chara/Thumb_004.png",
            asset_type="Sprite"
        )
        audio = self.route("Sound/AssetBundles/Audio/BGM/theme.wav", asset_type="AudioClip")
        self.assertEqual(sprite.as_posix(), "Sprites/Thumb_004.png")
        self.assertEqual(audio.as_posix(), "Audio/theme.wav")

    def test_sd_tf_st_are_inside_feature_and_costume_archive(self):
        expected_prefixes = ("sd", "tf", "st")
        for prefix in expected_prefixes:
            with self.subTest(prefix=prefix):
                actual = self.route(
                    f"{prefix}_1043002_SkeletonData.skel",
                    asset_id="1043002", bundle=f"{prefix}_1043002_SkeletonData_HighQuality.bundle",
                    asset_type="TextAsset", variant="HighQuality",
                    asset_root=f"Assets/Res/Spine/HighQuality/{prefix}_1043002_SkeletonData",
                    category="Res",
                )
                self.assertEqual(
                    actual.as_posix(),
                    f"104300/1043002/{prefix}_1043002/{prefix}_1043002_SkeletonData.skel",
                )

    def test_prefixed_skeleton_container_folder_is_not_preserved_as_directory(self):
        actual = self.route(
            "Assets/Res/Spine/HighQuality/st_1043002_SkeletonData/SkeletonData.skel",
            asset_id="1043002", bundle="st_1043002_SkeletonData_HighQuality.bundle",
            asset_type="TextAsset", variant="HighQuality",
            asset_root="Assets/Res/Spine/HighQuality/st_1043002_SkeletonData", category="Res",
        )
        self.assertEqual(actual.as_posix(), "104300/1043002/st_1043002/SkeletonData.skel")
        self.assertNotIn("st_1043002_SkeletonData/", actual.as_posix())

    def test_sd_tf_st_related_files_remain_in_same_character_quality_folder(self):
        names = [
            "sd_1043002_SkeletonData.skel",
            "sd_1043002_SkeletonData.atlas",
            "st_1043002_SkeletonData.png",
            "tf_1043002_SkeletonData.skel",
        ]
        outputs = [
            self.route(
                name, asset_id="1043002", bundle=f"{name}_HighQuality.bundle",
                asset_type="TextAsset", variant="HighQuality",
                asset_root="Assets/Res/HighQuality", category="Spine",
            )
            for name in names
        ]
        self.assertEqual({p.parent.as_posix() for p in outputs}, {"104300/1043002/sd_1043002", "104300/1043002/st_1043002", "104300/1043002/tf_1043002"})
        self.assertEqual([p.name for p in outputs], names)

    def test_generic_bundle_still_uses_object_prefix_and_quality_from_root(self):
        actual = self.route(
            "Assets/Res/st_1043002_SkeletonData/SkeletonData.skel",
            asset_id="1043002", bundle="generic_bundle.bundle",
            asset_type="TextAsset", variant="", asset_root="Assets/Res/HighQuality",
            category="Res",
        )
        self.assertEqual(actual.as_posix(), "104300/1043002/st_1043002/SkeletonData.skel")

    def test_prefix_skeletondata_name_embedded_in_catalog_id_is_recognized(self):
        actual = self.route(
            "Assets/Res/Spine/HighQuality/SkeletonData.skel",
            asset_id="st_1043002_SkeletonData", bundle="generic_spine.bundle",
            asset_type="TextAsset", variant="HighQuality",
            asset_root="Assets/Res/Spine/HighQuality", category="Res",
        )
        self.assertEqual(actual.as_posix(), "104300/1043002/st_1043002/SkeletonData.skel")

    def test_high_and_low_quality_separate_inside_same_character_folder(self):
        rows = [
            {"id": "1043002", "bundle": "sd_1043002_HighQuality.bundle", "version": "HighQuality", "category": "Res", "asset_root": "Assets/Res/HighQuality"},
            {"id": "1043002", "bundle": "st_1043002_LowQuality.bundle", "version": "LowQuality", "category": "Res", "asset_root": "Assets/Res/LowQuality"},
        ]
        variants = CORE.build_variant_subdirs(rows)
        self.assertEqual(variants["sd_1043002_highquality.bundle"], "HighQuality")
        self.assertEqual(variants["st_1043002_lowquality.bundle"], "LowQuality")
        high = self.route("sd_1043002_SkeletonData.skel", asset_id="1043002", bundle=rows[0]["bundle"], variant=variants[rows[0]["bundle"].casefold()], asset_root=rows[0]["asset_root"], category="Res")
        low = self.route("st_1043002_SkeletonData.skel", asset_id="1043002", bundle=rows[1]["bundle"], variant=variants[rows[1]["bundle"].casefold()], asset_root=rows[1]["asset_root"], category="Res")
        self.assertEqual(high.as_posix(), "104300/1043002/sd_1043002/sd_1043002_SkeletonData.skel")
        self.assertEqual(low.as_posix(), "104300/1043002/st_1043002/LowQuality/st_1043002_SkeletonData.skel")

    def test_sd_and_st_of_same_costume_share_variant_group_not_prefix_folder(self):
        rows = [
            {"id": "1043002", "bundle": "sd_1043002_SkeletonData.bundle", "version": "HighQuality", "category": "Spine", "asset_root": "Assets/Res/Spine/HighQuality/sd_1043002_SkeletonData"},
            {"id": "1043002", "bundle": "st_1043002_SkeletonData.bundle", "version": "HighQuality_", "category": "Spine", "asset_root": "Assets/Res/Spine/HighQuality_/st_1043002_SkeletonData"},
        ]
        self.assertEqual(CORE._variant_group_key(rows[0]), "character:1043002")
        self.assertEqual(CORE._variant_group_key(rows[1]), "character:1043002")
        variants = CORE.build_variant_subdirs(rows)
        self.assertEqual(variants["sd_1043002_skeletondata.bundle"], "HighQuality")
        self.assertEqual(variants["st_1043002_skeletondata.bundle"], "HighQuality")

    def test_tf_m0_file_goes_inside_canonical_tf_family_folder(self):
        actual = self.route(
            "tf_1043001_m0.skel", asset_id="1043001", bundle="tf_1043001_HighQuality.bundle",
            asset_type="TextAsset", variant="HighQuality", asset_root="Assets/Res/HighQuality/tf_1043001",
            category="Spine",
        )
        self.assertEqual(actual.as_posix(), "104300/1043001/tf_1043001/tf_1043001_m0.skel")

    def test_low_quality_only_adds_lowquality_folder(self):
        high = self.route("tf_1043001_m0.skel", asset_id="1043001", bundle="tf_1043001_HighQuality.bundle",
                          asset_type="TextAsset", variant="HighQuality", asset_root="Assets/Res/HighQuality", category="Spine")
        low = self.route("tf_1043001_m0.skel", asset_id="1043001", bundle="tf_1043001_LowQuality.bundle",
                         asset_type="TextAsset", variant="LowQuality", asset_root="Assets/Res/LowQuality", category="Spine")
        self.assertEqual(high.as_posix(), "104300/1043001/tf_1043001/tf_1043001_m0.skel")
        self.assertEqual(low.as_posix(), "104300/1043001/tf_1043001/LowQuality/tf_1043001_m0.skel")

    def test_adults_and_general_versions_are_kept_separate(self):
        adults = self.route("tf_1043001_m0.skel", asset_id="1043001", bundle="tf_1043001_Adults.bundle",
                            asset_type="TextAsset", variant="Adults", category="Spine")
        general = self.route("tf_1043001_m0.skel", asset_id="1043001", bundle="tf_1043001_General.bundle",
                             asset_type="TextAsset", variant="General", category="Spine")
        self.assertEqual(adults.as_posix(), "104300/1043001/tf_1043001/Adults/tf_1043001_m0.skel")
        self.assertEqual(general.as_posix(), "104300/1043001/tf_1043001/General/tf_1043001_m0.skel")

    def test_id_layout_matches_legacy_catalog_id_behavior(self):
        actual = self.route(
            "sd_1043002_SkeletonData.skel", asset_id="1043002",
            bundle="sd_1043002_SkeletonData_HighQuality.bundle",
            layout="id", variant="HighQuality",
        )
        self.assertEqual(actual.as_posix(), "1043002/HighQuality/sd_1043002_SkeletonData.skel")

    def test_original_layout_is_unchanged(self):
        path = PurePosixPath("Assets/Res/Spine/ch_100021.skel")
        self.assertEqual(self.route(path.as_posix(), asset_id="ch_100021", layout="original"), path)

    def test_high_quality_child_ch_folder_is_directly_under_costume_id(self):
        actual = self.route(
            "HighQuality/ch_1043001/body.prefab", asset_id="1043001",
            bundle="chara_1043001_HighQuality.bundle", asset_type="GameObject",
            variant="HighQuality", asset_root="Assets/Res/HighQuality/ch_1043001",
        )
        self.assertEqual(actual.as_posix(), "104300/1043001/ch_1043001/body.prefab")

    def test_high_quality_plain_costume_folder_is_directly_under_costume_id(self):
        actual = self.route(
            "HighQuality/1043001/1043001.prefab", asset_id="1043001",
            bundle="chara_1043001_HighQuality.bundle", asset_type="GameObject",
            variant="HighQuality", asset_root="Assets/Res/HighQuality/1043001",
        )
        self.assertEqual(actual.as_posix(), "104300/1043001/1043001/1043001.prefab")

    def test_low_quality_folder_is_inside_existing_child_family_folder(self):
        ch = self.route(
            "LowQuality/ch_1043001/body.prefab", asset_id="1043001",
            bundle="chara_1043001_LowQuality.bundle", asset_type="GameObject",
            variant="LowQuality", asset_root="Assets/Res/LowQuality/ch_1043001",
        )
        numeric = self.route(
            "LowQuality/1043001/1043001.prefab", asset_id="1043001",
            bundle="chara_1043001_LowQuality.bundle", asset_type="GameObject",
            variant="LowQuality", asset_root="Assets/Res/LowQuality/1043001",
        )
        self.assertEqual(ch.as_posix(), "104300/1043001/ch_1043001/LowQuality/body.prefab")
        self.assertEqual(numeric.as_posix(), "104300/1043001/1043001/LowQuality/1043001.prefab")

    def test_six_digit_ch_spine_retains_legacy_ch_id_folder(self):
        actual = self.route(
            "Assets/Res/Spine/HighQuality/ch_100021.skel",
            asset_id="ch_100021", bundle="ch_100021.bundle", asset_type="TextAsset",
            variant="HighQuality", asset_root="Assets/Res/Spine/HighQuality/ch_100021",
            category="Spine",
        )
        self.assertEqual(actual.as_posix(), "ch_100021/ch_100021.skel")

    def test_no_seven_digit_character_id_sd_st_assets_do_not_create_per_object_folders(self):
        for prefix in ("sd", "tf", "st"):
            with self.subTest(prefix=prefix):
                actual = self.route(
                    f"{prefix}_123_SkeletonData.skel", asset_id="", bundle="shared_spine.bundle",
                    asset_type="TextAsset", category="Res",
                )
                self.assertEqual(actual.as_posix(), f"Spine/{prefix}_123_SkeletonData.skel")


class VariantPlanningTests(unittest.TestCase):
    def test_high_quality_trailing_underscore_normalized(self):
        rows = [
            {"id": "1043002", "bundle": "sd_1043002_a.bundle", "version": "HighQuality", "category": "Spine", "asset_root": "Assets/Res/Spine/HighQuality/sd_1043002"},
            {"id": "1043002", "bundle": "st_1043002_b.bundle", "version": "HighQuality_", "category": "Spine", "asset_root": "Assets/Res/Spine/HighQuality_/st_1043002"},
        ]
        actual = CORE.build_variant_subdirs(rows)
        self.assertEqual(actual["sd_1043002_a.bundle"], "HighQuality")
        self.assertEqual(actual["st_1043002_b.bundle"], "HighQuality")

    def test_quality_is_inferred_from_asset_root_or_bundle_when_version_empty(self):
        rows = [
            {"id": "1043002", "bundle": "sd_1043002.bundle", "version": "", "category": "Res", "asset_root": "Assets/Res/HighQuality/sd_1043002"},
            {"id": "1043002", "bundle": "st_1043002_lq.bundle", "version": "", "category": "Res", "asset_root": "Assets/Res/LowQuality/st_1043002"},
        ]
        actual = CORE.build_variant_subdirs(rows)
        self.assertEqual(actual["sd_1043002.bundle"], "HighQuality")
        self.assertEqual(actual["st_1043002_lq.bundle"], "LowQuality")

    def test_plain_numeric_single_version_preserves_legacy_no_version_folder(self):
        rows = [{"id": "1043002", "bundle": "character.bundle", "version": "HighQuality", "category": "Textures", "asset_root": "Assets/Res/Textures"}]
        # Ordinary numeric IDs keep legacy semantics; only explicit SD/TF/ST Spine
        # bundles force a known quality folder when a single quality is selected.
        actual = CORE.build_variant_subdirs(rows)
        self.assertEqual(actual["character.bundle"], "")

    def test_explicit_id_layout_retains_v7_catalog_id_version_planning(self):
        one_version = [{"id": "sd_1043002_SkeletonData", "bundle": "sd_hq.bundle", "version": "HighQuality", "category": "Spine", "asset_root": "Assets/Res/Spine/HighQuality"}]
        self.assertEqual(CORE.build_variant_subdirs(one_version, "id")["sd_hq.bundle"], "")
        two_versions = [
            {"id": "sd_1043002_SkeletonData", "bundle": "sd_hq.bundle", "version": "HighQuality", "category": "Spine", "asset_root": "Assets/Res/Spine/HighQuality"},
            {"id": "sd_1043002_SkeletonData", "bundle": "sd_lq.bundle", "version": "LowQuality", "category": "Spine", "asset_root": "Assets/Res/Spine/LowQuality"},
        ]
        planned = CORE.build_variant_subdirs(two_versions, "id")
        self.assertEqual(planned["sd_hq.bundle"], "HighQuality")
        self.assertEqual(planned["sd_lq.bundle"], "LowQuality")


class StaleExportCleanupTests(unittest.TestCase):
    def _write_manifest(self, path, rows):
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=CORE.EXPORT_MANIFEST_FIELDS)
            writer.writeheader()
            writer.writerows(rows)

    def test_removes_old_manifest_paths_only_for_successfully_reexported_bundle(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            export_root = root / "exported_assets"
            logs_root = root / "logs"
            old_dir = export_root / "104300/1043002" / "HighQuality_"
            old_dir.mkdir(parents=True)
            old_file = old_dir / "SkeletonData.skel"
            old_file.write_bytes(b"old")
            keep_file = export_root / "Audio" / "still-used.wav"
            keep_file.parent.mkdir(parents=True)
            keep_file.write_bytes(b"keep")
            old_manifest = logs_root / "20261010" / "round_001_export_manifest.csv"
            self._write_manifest(old_manifest, [
                {"bundle": "st_1043002_SkeletonData_HighQuality.bundle", "asset_path": "Assets/Res/Spine/st_1043002_SkeletonData/SkeletonData.skel", "exported_path": "104300/1043002/HighQuality_/SkeletonData.skel", "type": "TextAsset", "path_id": "1", "mode": "textasset-skel", "status": "ok"},
                {"bundle": "other.bundle", "asset_path": "Assets/Audio/theme", "exported_path": "Audio/still-used.wav", "type": "AudioClip", "path_id": "2", "mode": "audio", "status": "ok"},
            ])
            current_manifest = logs_root / "20261012" / "round_001_export_manifest.csv"
            self._write_manifest(current_manifest, [
                {"bundle": "st_1043002_SkeletonData_HighQuality.bundle", "asset_path": "Assets/Res/Spine/st_1043002_SkeletonData/SkeletonData.skel", "exported_path": "104300/1043002/HighQuality/SkeletonData.skel", "type": "TextAsset", "path_id": "1", "mode": "textasset-skel", "status": "ok"},
            ])
            new_file = export_root / "104300/1043002/HighQuality" / "SkeletonData.skel"
            new_file.parent.mkdir(parents=True)
            new_file.write_bytes(b"new")
            removed_files, removed_dirs = CORE.reconcile_stale_export_paths(
                export_root, logs_root, current_manifest, {"st_1043002_SkeletonData_HighQuality.bundle"}, None,
            )
            self.assertEqual(removed_files, 1)
            self.assertFalse(old_file.exists())
            self.assertTrue(new_file.exists())
            self.assertTrue(keep_file.exists())

    def test_manifest_path_traversal_is_never_deleted(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            export_root = root / "exported_assets"
            logs_root = root / "logs"
            outside = root / "outside.skel"
            outside.write_bytes(b"do not delete")
            old_manifest = logs_root / "old" / "round_001_export_manifest.csv"
            self._write_manifest(old_manifest, [
                {"bundle": "st_1043002.bundle", "asset_path": "Assets/Res/st_1043002", "exported_path": "../outside.skel", "type": "TextAsset", "path_id": "1", "mode": "raw", "status": "ok"},
            ])
            current_manifest = logs_root / "new" / "round_001_export_manifest.csv"
            self._write_manifest(current_manifest, [])
            CORE.reconcile_stale_export_paths(export_root, logs_root, current_manifest, {"st_1043002.bundle"}, None)
            self.assertTrue(outside.exists())

    def test_repair_existing_exports_moves_legacy_prefix_folder_into_character_archive(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            export_root = root / "exported_assets"
            logs_root = root / "logs"
            old_file = export_root / "st_1043002_SkeletonData" / "HighQuality_" / "SkeletonData.skel"
            old_file.parent.mkdir(parents=True)
            old_file.write_bytes(b"spine-data")
            untracked = export_root / "untracked" / "custom.bin"
            untracked.parent.mkdir(parents=True)
            untracked.write_bytes(b"keep-me")
            manifest = logs_root / "20261010" / "round_001_export_manifest.csv"
            self._write_manifest(manifest, [
                {"bundle": "st_1043002_SkeletonData_HighQuality_.bundle", "asset_path": "Assets/Res/Spine/st_1043002_SkeletonData/SkeletonData.skel", "exported_path": "st_1043002_SkeletonData/HighQuality_/SkeletonData.skel", "type": "TextAsset", "path_id": "10", "mode": "textasset-skel", "status": "ok"},
            ])
            rows = [{
                "bundle": "st_1043002_SkeletonData_HighQuality_.bundle", "id": "1043002", "version": "",
                "category": "Res", "asset_root": "Assets/Res/Spine/HighQuality/st_1043002_SkeletonData", "asset_count": "1",
            }]
            result = CORE.repair_existing_exports(export_root, logs_root, rows, "character")
            new_file = export_root / "104300" / "1043002" / "st_1043002" / "SkeletonData.skel"
            self.assertEqual(result["moved"], 1)
            self.assertFalse(old_file.exists())
            self.assertTrue(new_file.is_file())
            self.assertTrue(untracked.is_file())

    def test_repair_flattens_highquality_wrapper_and_recovers_physical_ch_folder(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            export_root = root / "old_project" / "exported_assets"
            # Simulate launching the newer GUI from another folder while the
            # old game's exported_assets and logs remain together in v1.6.
            logs_root = root / "new_project_logs"
            old_logs_root = export_root.parent / "logs"
            old_file = export_root / "104300" / "1043001" / "HighQuality" / "ch_1043001" / "body.prefab"
            old_file.parent.mkdir(parents=True)
            old_file.write_bytes(b"body-prefab")
            manifest = old_logs_root / "20261012" / "round_001_export_manifest.csv"
            self._write_manifest(manifest, [
                {"bundle": "chara_1043001_HighQuality.bundle", "asset_path": "Assets/Res/Prefabs/body.prefab", "exported_path": "104300/1043001/HighQuality/ch_1043001/body.prefab", "type": "GameObject", "path_id": "1", "mode": "prefab", "status": "ok"},
            ])
            rows = [{
                "bundle": "chara_1043001_HighQuality.bundle", "id": "1043001", "version": "HighQuality",
                "category": "Prefabs", "asset_root": "Assets/Res/HighQuality", "asset_count": "1",
            }]
            result = CORE.repair_existing_exports(export_root, logs_root, rows, "character")
            target = export_root / "104300" / "1043001" / "ch_1043001" / "body.prefab"
            self.assertEqual(result["moved"], 1)
            self.assertFalse(old_file.exists())
            self.assertTrue(target.is_file())
            self.assertFalse((export_root / "104300" / "1043001" / "HighQuality").exists())

    def test_repair_flattens_highquality_wrapper_and_recovers_numeric_child_folder(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            export_root = root / "exported_assets"
            logs_root = root / "logs"
            old_file = export_root / "104300" / "1043001" / "HighQuality" / "1043001" / "body.prefab"
            old_file.parent.mkdir(parents=True)
            old_file.write_bytes(b"numeric-body-prefab")
            manifest = logs_root / "20261012" / "round_001_export_manifest.csv"
            self._write_manifest(manifest, [
                {"bundle": "numeric_1043001_HighQuality.bundle", "asset_path": "Assets/Res/Prefabs/body.prefab", "exported_path": "104300/1043001/HighQuality/1043001/body.prefab", "type": "GameObject", "path_id": "1", "mode": "prefab", "status": "ok"},
            ])
            rows = [{
                "bundle": "numeric_1043001_HighQuality.bundle", "id": "1043001", "version": "HighQuality",
                "category": "Prefabs", "asset_root": "Assets/Res/HighQuality", "asset_count": "1",
            }]
            result = CORE.repair_existing_exports(export_root, logs_root, rows, "character")
            target = export_root / "104300" / "1043001" / "1043001" / "body.prefab"
            self.assertEqual(result["moved"], 1)
            self.assertFalse(old_file.exists())
            self.assertTrue(target.is_file())
            self.assertFalse((export_root / "104300" / "1043001" / "HighQuality").exists())

    def test_repair_writes_new_manifest_for_future_cleanup_tracking(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            export_root = root / "exported_assets"
            logs_root = root / "logs"
            old_file = export_root / "st_1043002_SkeletonData" / "HighQuality_" / "SkeletonData.skel"
            old_file.parent.mkdir(parents=True)
            old_file.write_bytes(b"spine-data")
            old_manifest = logs_root / "old" / "round_001_export_manifest.csv"
            self._write_manifest(old_manifest, [
                {"bundle": "st_1043002_SkeletonData_HighQuality_.bundle", "asset_path": "Assets/Res/Spine/st_1043002_SkeletonData/SkeletonData.skel", "exported_path": "st_1043002_SkeletonData/HighQuality_/SkeletonData.skel", "type": "TextAsset", "path_id": "10", "mode": "textasset-skel", "status": "ok"},
            ])
            rows = [{
                "bundle": "st_1043002_SkeletonData_HighQuality_.bundle", "id": "1043002", "version": "",
                "category": "Res", "asset_root": "Assets/Res/Spine/HighQuality/st_1043002_SkeletonData", "asset_count": "1",
            }]
            session = CORE.SessionLog(logs_root)
            CORE.repair_existing_exports(export_root, logs_root, rows, "character", session)
            repair_manifest = session.root / "repair_export_manifest.csv"
            self.assertTrue(repair_manifest.is_file())
            with repair_manifest.open("r", encoding="utf-8-sig", newline="") as f:
                records = list(csv.DictReader(f))
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["exported_path"], "104300/1043002/st_1043002/SkeletonData.skel")
            self.assertEqual(records[0]["status"], "ok")


if __name__ == "__main__":
    unittest.main(verbosity=2)
