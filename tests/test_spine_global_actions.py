import importlib.util
from pathlib import Path
import unittest

MODULE_PATH = Path(__file__).resolve().parents[1] / "core" / "spine_global_actions.py"
spec = importlib.util.spec_from_file_location("spine_global_actions_under_test", MODULE_PATH)
spine = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(spine)


class SpineActionKeyframeTests(unittest.TestCase):
    def setUp(self):
        self.data = {"bones": [{"name": "root"}], "animations": {}}

    def test_manual_keyframes_are_source_of_truth_for_preset_type(self):
        segments = [
            {"time": 0, "x": 0, "y": 0},
            {"time": 0.123, "x": 17, "y": -3},
            {"time": 0.62, "x": -8, "y": 11},
            {"time": 0.9, "x": 0, "y": 0},
        ]
        spine._inject(self.data, [{
            "name": "ManualJump", "kind": "跳跃", "dur": 0.9, "amp": 48,
            "count": 1, "repeat": 1, "direction": "上", "segments": segments,
        }], "root")
        actual = self.data["animations"]["ManualJump"]["bones"]["root"]["translate"]
        self.assertEqual(actual, segments)

    def test_legacy_preset_without_segments_still_generates_keys(self):
        spine._inject(self.data, [{
            "name": "LegacyJump", "kind": "跳跃", "dur": 0.4, "amp": 20,
            "count": 2, "repeat": 1, "direction": "上", "segments": [],
        }], "root")
        actual = self.data["animations"]["LegacyJump"]["bones"]["root"]["translate"]
        self.assertGreaterEqual(len(actual), 9)
        self.assertEqual(actual[0]["time"], 0)
        self.assertAlmostEqual(actual[-1]["time"], 0.4)
        self.assertEqual(actual[-1]["x"], 0)
        self.assertEqual(actual[-1]["y"], 0)

    def test_repeat_is_applied_to_saved_keyframes(self):
        segments = [
            {"time": 0, "x": 0, "y": 0},
            {"time": 0.2, "x": 12, "y": 4},
            {"time": 0.5, "x": 0, "y": 0},
        ]
        spine._inject(self.data, [{
            "name": "Repeated", "kind": "左右摇摆", "dur": 0.5, "amp": 22,
            "count": 1, "repeat": 2, "direction": "左", "segments": segments,
        }], "root")
        actual = self.data["animations"]["Repeated"]["bones"]["root"]["translate"]
        self.assertEqual(len(actual), 5)
        self.assertAlmostEqual(actual[-1]["time"], 1.0)
        self.assertEqual(actual[1]["x"], 12)
        self.assertEqual(actual[2]["x"], 0)
        self.assertAlmostEqual(actual[3]["time"], 0.7)

    def test_duplicate_animation_name_uses_suffix_without_overwriting(self):
        self.data["animations"]["Jump"] = {"bones": {}}
        spine._inject(self.data, [{
            "name": "Jump", "kind": "跳跃", "dur": 0.2, "amp": 48,
            "count": 1, "repeat": 1, "direction": "上", "segments": [],
        }], "root")
        self.assertIn("Jump", self.data["animations"])
        self.assertIn("Jump_EX2", self.data["animations"])

    def test_non_increasing_keyframe_time_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "严格递增"):
            spine._inject(self.data, [{
                "name": "Bad", "kind": "自定义", "dur": 0.2, "amp": 48,
                "count": 1, "repeat": 1, "direction": "上",
                "segments": [{"time": 0, "x": 0, "y": 0}, {"time": 0, "x": 5, "y": 0}],
            }], "root")

    def test_out_of_order_keyframes_are_not_silently_sorted(self):
        with self.assertRaisesRegex(ValueError, "严格递增"):
            spine._inject(self.data, [{
                "name": "OutOfOrder", "kind": "自定义", "dur": 0.5, "amp": 48,
                "count": 1, "repeat": 1, "direction": "上",
                "segments": [
                    {"time": 0, "x": 0, "y": 0},
                    {"time": 0.5, "x": 8, "y": 0},
                    {"time": 0.25, "x": 4, "y": 0},
                ],
            }], "root")

    def test_non_finite_keyframe_values_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "无效数值"):
            spine._inject(self.data, [{
                "name": "NotFinite", "kind": "自定义", "dur": 0.5, "amp": 48,
                "count": 1, "repeat": 1, "direction": "上",
                "segments": [
                    {"time": 0, "x": 0, "y": 0},
                    {"time": 0.5, "x": float("inf"), "y": 0},
                ],
            }], "root")

    def test_filter_stays_limited_to_explicit_sd_ch_skel(self):
        self.assertTrue(spine.is_sd_ch_spine(r"C:\assets\sd_10001.skel"))
        self.assertTrue(spine.is_sd_ch_spine(r"C:\assets\ch_10001\skeleton.skel"))
        self.assertFalse(spine.is_sd_ch_spine(r"C:\assets\tf_10001.skel"))
        self.assertFalse(spine.is_sd_ch_spine(r"C:\assets\chara_10001.skel"))


class ActionTypeTemplateSyncTests(unittest.TestCase):
    def test_changing_type_imports_matching_template(self):
        original = {
            "name": "NewAction", "kind": "跳跃", "dur": 0.4, "amp": 30,
            "count": 2, "repeat": 1, "direction": "右",
            "segments": [{"time": 0, "x": 0, "y": 0}, {"time": 0.4, "x": 0, "y": 0}],
        }
        changed = spine.sync_action_type_template({**original, "kind": "左右摇摆"}, "跳跃")
        self.assertGreaterEqual(len(changed["segments"]), 5)
        self.assertAlmostEqual(changed["segments"][-1]["time"], 0.4)
        self.assertTrue(any(abs(key["x"]) > 0 for key in changed["segments"]))
        self.assertTrue(all(key["y"] == 0 for key in changed["segments"]))
        self.assertEqual(changed["kind"], "左右摇摆")

    def test_changing_to_custom_preserves_current_keys(self):
        keys = [{"time": 0, "x": 0, "y": 0}, {"time": 0.17, "x": 23, "y": -5}]
        action = {"kind": "自定义", "dur": 0.17, "amp": 48, "count": 1,
                  "direction": "上", "segments": keys}
        changed = spine.sync_action_type_template(action, "跳跃")
        self.assertEqual(changed["segments"], keys)

    def test_same_type_does_not_overwrite_manual_edits(self):
        keys = [{"time": 0, "x": 0, "y": 0}, {"time": 0.123, "x": 9, "y": 13}]
        action = {"kind": "跳跃", "dur": 0.123, "amp": 48, "count": 1,
                  "direction": "上", "segments": keys}
        changed = spine.sync_action_type_template(action, "跳跃")
        self.assertEqual(changed["segments"], keys)

    def test_type_change_uses_current_amplitude_count_direction(self):
        action = {"kind": "左右摇摆", "dur": 0.3, "amp": 17, "count": 2,
                  "direction": "右", "segments": []}
        changed = spine.sync_action_type_template(action, "自定义")
        self.assertAlmostEqual(changed["segments"][1]["x"], 17)
        self.assertEqual(changed["segments"][-1]["time"], 0.3)
        self.assertGreaterEqual(len(changed["segments"]), 9)

    def test_invalid_template_parameters_are_rejected(self):
        action = {"kind": "跳跃", "dur": 0, "amp": 48, "count": 1,
                  "direction": "上", "segments": []}
        with self.assertRaisesRegex(ValueError, "时长"):
            spine.sync_action_type_template(action, "自定义")


if __name__ == "__main__":
    unittest.main()
