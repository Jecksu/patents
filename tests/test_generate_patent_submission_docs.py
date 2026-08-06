from __future__ import annotations

import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from PIL import Image


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "generate_patent_submission_docs.py"
SPEC = importlib.util.spec_from_file_location("generate_patent_submission_docs", SCRIPT)
generator = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(generator)


class PatentSubmissionMaterialTests(unittest.TestCase):
    def setUp(self) -> None:
        self.lightweight = next(
            material for material in generator.MATERIALS if material["key"] == "lightweight_strengthened"
        )

    def test_strengthened_lightweight_material_uses_independent_key(self) -> None:
        self.assertEqual(self.lightweight["key"], "lightweight_strengthened")

    def test_strengthened_lightweight_material_uses_strengthened_source(self) -> None:
        self.assertEqual(
            self.lightweight["source"],
            generator.PATENT_DIR
            / "2026-08-06-cad-model-diff-lightweight-summary-strengthened-patent-draft.md",
        )

    def test_strengthened_lightweight_material_uses_strengthened_output(self) -> None:
        self.assertEqual(
            self.lightweight["output"],
            generator.OUT_DIR
            / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件-强化版.docx",
        )

    def test_strengthened_lightweight_title_matches_patent_draft(self) -> None:
        self.assertEqual(
            self.lightweight["title"],
            "一种面向跨CAD内核的基于轻量摘要闸门和受限局部BREP验证的CAD模型差异对比方法及系统",
        )

    def test_strengthened_lightweight_has_five_matching_figure_titles(self) -> None:
        self.assertEqual(
            self.lightweight["figure_titles"],
            [
                "图1 方法闭环流程示意图",
                "图2 系统模块与比较会话数据流示意图",
                "图3 多变换假设与摘要闸门示意图",
                "图4 候选关系消歧、局部验证和差异闭包示意图",
                "图5 场景处理与降级示意图",
            ],
        )

    def test_method_figure_covers_the_strengthened_comparison_loop(self) -> None:
        labels = " ".join(self.lightweight["abstract_figure"])
        for expected in (
            "无可比较永久命名",
            "至少两类轻量摘要",
            "多变换假设",
            "至少两个摘要闸门",
            "候选",
            "排除",
            "稀疏候选图",
            "互斥消歧",
            "预算受限局部BREP",
            "R_11/R_1m/R_m1",
            "REVIEW保护",
            "关系闭包",
            "语义聚合",
            "双侧定位",
        ):
            self.assertIn(expected, labels)

    def test_system_figure_modules_match_the_strengthened_system(self) -> None:
        labels = " ".join(self.lightweight["system_modules"])
        for expected in (
            "审计入口",
            "轻量摘要提取",
            "变换假设",
            "摘要闸门",
            "候选关系消歧",
            "局部BREP按需读取",
            "关系/谱系",
            "差异闭包",
            "语义聚合",
            "双侧定位",
        ):
            self.assertIn(expected, labels)

    def test_decision_figure_covers_multiple_hypotheses_and_gate_outputs(self) -> None:
        labels = " ".join(self.lightweight["decision_flow"])
        for expected in (
            "刚体",
            "单位",
            "镜像",
            "仿射",
            "对称",
            "验证并保留多假设",
            "至少两个摘要闸门",
            "候选关系",
            "排除关系",
            "全局差异不默认终止",
        ):
            self.assertIn(expected, labels)

    def test_materials_use_one_ordered_figure_protocol_and_explicit_layouts(self) -> None:
        permanent = next(material for material in generator.MATERIALS if material["key"] == "permanent_naming")
        self.assertEqual(
            self.lightweight["figure_kinds"],
            ("method", "system", "decision", "candidate_closure", "scenarios"),
        )
        self.assertEqual(self.lightweight["layout"], "strengthened")
        self.assertEqual(
            permanent["figure_kinds"],
            ("method", "system", "decision", "metrics_diff"),
        )
        self.assertEqual(permanent["layout"], "standard")
        for material in generator.MATERIALS:
            self.assertNotIn("extra_figure", material)
            self.assertNotIn("extra_figures", material)

    def test_candidate_closure_and_scenario_figures_are_nonempty_pngs(self) -> None:
        self.assertTrue(hasattr(generator, "make_candidate_closure_flow"))
        self.assertTrue(hasattr(generator, "make_scenario_flow"))
        with tempfile.TemporaryDirectory() as temp_dir:
            candidate_path = Path(temp_dir) / "candidate.png"
            scenario_path = Path(temp_dir) / "scenario.png"
            generator.make_candidate_closure_flow(candidate_path, "图4")
            generator.make_scenario_flow(scenario_path, "图5")
            for path in (candidate_path, scenario_path):
                self.assertTrue(path.is_file())
                self.assertGreater(path.stat().st_size, 0)
                self.assertEqual(path.read_bytes()[:8], b"\x89PNG\r\n\x1a\n")

    def test_candidate_closure_labels_expose_all_domain_steps(self) -> None:
        self.assertTrue(hasattr(generator, "CANDIDATE_CLOSURE_LABELS"))
        labels = generator.CANDIDATE_CLOSURE_LABELS
        expected_by_node = (
            ("C_gate",),
            ("候选裕量", "双向排名", "端点互斥"),
            ("预算受限局部BREP",),
            ("R_11/R_1m/R_m1",),
            ("REVIEW保护",),
            ("COMMON_11", "保持不变", "修改", "DEL", "ADD"),
            ("设计语义", "双侧定位"),
        )
        self.assertEqual(len(labels), len(expected_by_node))
        for label, expected_tokens in zip(labels, expected_by_node, strict=True):
            for expected in expected_tokens:
                self.assertIn(expected, label)

    def test_scenario_rows_expose_all_processing_and_degradation_paths(self) -> None:
        self.assertTrue(hasattr(generator, "SCENARIO_ROWS"))
        expected_rows = (
            ("刚体/单位", "归一化继续"),
            ("未知仿射/镜像", "全局差异继续局部"),
            ("对称多假设", "消歧或REVIEW"),
            ("分裂/合并", "组关系验证"),
            ("周期缝边/解析-NURBS", "等价表达验证"),
            ("数据缺失/读取失败/无效拓扑", "REVIEW或受控回退"),
            ("整体不相关", "全局证据+分区排除"),
        )
        self.assertEqual(len(generator.SCENARIO_ROWS), len(expected_rows))
        for (scenario, action), (expected_scenario, expected_action) in zip(
            generator.SCENARIO_ROWS, expected_rows, strict=True
        ):
            self.assertIn(expected_scenario, scenario)
            self.assertIn(expected_action, action)

    def test_permanent_naming_material_and_output_are_preserved(self) -> None:
        permanent = next(material for material in generator.MATERIALS if material["key"] == "permanent_naming")
        self.assertEqual(
            permanent["output"],
            generator.OUT_DIR / "永久命名-CAD模型差异对比-发明专利申请文件.docx",
        )
        self.assertEqual(permanent["figure_kinds"][3], "metrics_diff")

    def test_old_lightweight_material_and_output_are_not_registered(self) -> None:
        self.assertNotIn("lightweight", {material["key"] for material in generator.MATERIALS})
        old_output = generator.OUT_DIR / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件.docx"
        self.assertNotIn(old_output, {material["output"] for material in generator.MATERIALS})

    def test_build_strengthened_material_generates_five_assets_and_six_drawings(self) -> None:
        material = dict(self.lightweight)
        expected_assets = {
            "lightweight_strengthened-fig1-method.png",
            "lightweight_strengthened-fig2-system.png",
            "lightweight_strengthened-fig3-decision.png",
            "lightweight_strengthened-fig4-candidate-closure.png",
            "lightweight_strengthened-fig5-scenarios.png",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            asset_dir = temp_root / "assets"
            output = temp_root / "strengthened.docx"
            material["output"] = output
            with mock.patch.object(generator, "ASSET_DIR", asset_dir):
                generator.build_material(material)

            self.assertEqual({path.name for path in asset_dir.iterdir()}, expected_assets)
            for asset_name in expected_assets:
                self.assertGreater((asset_dir / asset_name).stat().st_size, 0)
            with Image.open(asset_dir / "lightweight_strengthened-fig1-method.png") as image:
                self.assertEqual(image.size, (1500, 2060))
            with Image.open(asset_dir / "lightweight_strengthened-fig3-decision.png") as image:
                self.assertEqual(image.size, (1900, 1350))
            drawing_count, media_count = self._docx_drawing_and_media_counts(output)
            self.assertEqual(drawing_count, 6)
            self.assertEqual(media_count, 5)
            self.assertFalse(
                (
                    temp_root
                    / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件.docx"
                ).exists()
            )

    def test_build_permanent_material_uses_metrics_branch_and_four_assets(self) -> None:
        permanent = dict(next(material for material in generator.MATERIALS if material["key"] == "permanent_naming"))
        expected_assets = {
            "permanent_naming-fig1-method.png",
            "permanent_naming-fig2-system.png",
            "permanent_naming-fig3-decision.png",
            "permanent_naming-fig4-metrics-diff.png",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            asset_dir = temp_root / "assets"
            output = temp_root / "permanent.docx"
            permanent["output"] = output
            with mock.patch.object(generator, "ASSET_DIR", asset_dir):
                generator.build_material(permanent)

            self.assertEqual({path.name for path in asset_dir.iterdir()}, expected_assets)
            for asset_name in expected_assets:
                self.assertGreater((asset_dir / asset_name).stat().st_size, 0)
            drawing_count, media_count = self._docx_drawing_and_media_counts(output)
            self.assertEqual(drawing_count, 5)
            self.assertEqual(media_count, 4)

    def test_swapped_strengthened_figure_order_is_rejected_before_writing(self) -> None:
        kinds = list(self._expected_strengthened_kinds())
        kinds[3], kinds[4] = kinds[4], kinds[3]
        self._assert_invalid_material_does_not_write(
            {"figure_kinds": tuple(kinds)},
            "图号|顺序|figure",
        )

    def test_unknown_figure_kind_is_rejected_before_writing(self) -> None:
        self._assert_invalid_material_does_not_write(
            {"figure_kinds": ("method", "system", "decision", "unknown", "scenarios")},
            "未知|注册|unknown",
        )

    def test_duplicate_figure_kind_is_rejected_before_writing(self) -> None:
        self._assert_invalid_material_does_not_write(
            {"figure_kinds": ("method", "system", "decision", "decision", "scenarios")},
            "重复|decision",
        )

    def test_more_figure_titles_than_kinds_is_rejected_before_writing(self) -> None:
        self._assert_invalid_material_does_not_write(
            {"figure_titles": [*self.lightweight["figure_titles"], "图6 多余附图"]},
            "数量|figure_titles",
        )

    def test_fewer_figure_titles_than_kinds_is_rejected_before_writing(self) -> None:
        self._assert_invalid_material_does_not_write(
            {"figure_titles": self.lightweight["figure_titles"][:-1]},
            "数量|figure_titles",
        )

    def test_mismatched_figure_title_number_is_rejected_before_writing(self) -> None:
        titles = list(self.lightweight["figure_titles"])
        titles[3] = titles[3].replace("图4", "图5", 1)
        self._assert_invalid_material_does_not_write(
            {"figure_titles": titles},
            "图4|标题|title",
        )

    def test_unknown_layout_is_rejected_before_writing(self) -> None:
        self._assert_invalid_material_does_not_write(
            {"layout": "unknown"},
            "layout|布局|unknown",
        )

    @staticmethod
    def _expected_strengthened_kinds() -> tuple[str, ...]:
        return ("method", "system", "decision", "candidate_closure", "scenarios")

    def _assert_invalid_material_does_not_write(self, updates: dict, message_pattern: str) -> None:
        material = dict(self.lightweight)
        material.update(updates)
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            asset_dir = temp_root / "assets"
            output = temp_root / "invalid.docx"
            material["output"] = output
            with mock.patch.object(generator, "ASSET_DIR", asset_dir):
                with self.assertRaisesRegex(ValueError, message_pattern):
                    generator.build_material(material)
            self.assertFalse(asset_dir.exists())
            self.assertFalse(output.exists())

    @staticmethod
    def _docx_drawing_and_media_counts(path: Path) -> tuple[int, int]:
        with zipfile.ZipFile(path) as archive:
            document_xml = archive.read("word/document.xml")
            media = [
                name
                for name in archive.namelist()
                if name.startswith("word/media/") and not name.endswith("/")
            ]
        return document_xml.count(b"<w:drawing"), len(media)


if __name__ == "__main__":
    unittest.main()
