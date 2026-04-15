from __future__ import annotations

import importlib.util
import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from unittest.mock import patch


MODULE_PATH = (
    Path(__file__).resolve().parents[1] / "scripts" / "doc_to_slides.py"
)
SPEC = importlib.util.spec_from_file_location("doc_to_slides", MODULE_PATH)
doc_to_slides = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(doc_to_slides)


class SourceChoiceTests(unittest.TestCase):
    def test_choose_source_candidate_persists_selected_search_hit(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            resolved_source = Path(tmpdir) / "resolved-source.json"
            doc_to_slides.write_json(
                resolved_source,
                {
                    "input_kind": "doc_name",
                    "resolved_kind": "",
                    "resolved_value": "",
                    "title": "",
                    "needs_user_choice": True,
                    "search_candidates": [
                        {
                            "title": "Quarterly Review",
                            "resolved_kind": "doc_url",
                            "resolved_value": "https://example.feishu.cn/docx/doccn123",
                        },
                        {
                            "title": "Quarterly Review Wiki",
                            "resolved_kind": "wiki_url",
                            "resolved_value": "https://example.feishu.cn/wiki/wiki_token",
                        },
                    ],
                },
            )

            chosen = doc_to_slides.choose_source_candidate(resolved_source, 2)

        self.assertEqual("wiki_url", chosen["resolved_kind"])
        self.assertEqual("https://example.feishu.cn/wiki/wiki_token", chosen["resolved_value"])
        self.assertEqual("Quarterly Review Wiki", chosen["title"])
        self.assertFalse(chosen["needs_user_choice"])
        self.assertEqual(2, chosen["selected_candidate_index"])

    def test_choose_source_candidate_rejects_out_of_range_selection(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            resolved_source = Path(tmpdir) / "resolved-source.json"
            doc_to_slides.write_json(
                resolved_source,
                {
                    "input_kind": "doc_name",
                    "resolved_kind": "",
                    "resolved_value": "",
                    "title": "",
                    "needs_user_choice": True,
                    "search_candidates": [
                        {
                            "title": "Quarterly Review",
                            "resolved_kind": "doc_url",
                            "resolved_value": "https://example.feishu.cn/docx/doccn123",
                        }
                    ],
                },
            )

            with self.assertRaisesRegex(ValueError, r"candidate-index must be between 1 and 1"):
                doc_to_slides.choose_source_candidate(resolved_source, 2)


class OutlineContractTests(unittest.TestCase):
    def make_outline(self) -> dict:
        return {
            "presentation": {
                "title": "Theme Review",
                "target_mode": "new",
                "content_mode": "report",
                "source": {
                    "input_kind": "doc_url",
                    "resolved_kind": "doc_url",
                    "resolved_value": "https://example.feishu.cn/docs/doccn123",
                },
            },
            "slides": [
                {
                    "no": 1,
                    "role": "content",
                    "section_divider": False,
                    "title": "Status",
                    "layout": "title-body",
                    "key_points": ["Point A", "Point B"],
                },
                {
                    "no": 2,
                    "role": "content",
                    "section_divider": False,
                    "title": "Next",
                    "layout": "title-body",
                    "key_points": ["Point C"],
                },
            ],
        }

    def test_validate_outline_rejects_exceeding_max_slides(self) -> None:
        outline = self.make_outline()
        outline["presentation"]["max_slides"] = 1

        with self.assertRaisesRegex(ValueError, r"slides exceed presentation.max_slides: 2 > 1"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_invalid_section_divider_shape(self) -> None:
        outline = self.make_outline()
        outline["slides"][0] = {
            "no": 1,
            "role": "content",
            "section_divider": True,
            "title": "Highlights",
            "layout": "title-only",
            "key_points": [],
        }

        with self.assertRaisesRegex(
            ValueError,
            r"section divider must use role='section' and layout='title-only'",
        ):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_title_that_exceeds_layout_budget(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["title"] = "非常长的标题" * 10

        with self.assertRaisesRegex(ValueError, r"slide 1 title is too long"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_overlong_key_point_for_comparison_layout(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["layout"] = "comparison"
        outline["slides"][0]["key_points"] = ["这是一条明显会超过比较布局文本预算的超长要点" * 3]

        with self.assertRaisesRegex(ValueError, r"key_points\[1\] is too long"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_overlong_timeline_total_budget(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["layout"] = "timeline"
        outline["slides"][0]["key_points"] = [
            "这是一条较长的时间线说明" * 2,
            "这是一条较长的时间线说明" * 2,
            "这是一条较长的时间线说明" * 2,
            "这是一条较长的时间线说明" * 2,
        ]

        with self.assertRaisesRegex(ValueError, r"total key point budget exceeded"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_overlong_metrics_point(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["layout"] = "metrics"
        outline["slides"][0]["key_points"] = ["非常长的指标标签内容明显超过卡片可承载长度"]

        with self.assertRaisesRegex(ValueError, r"key_points\[1\] is too long"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_non_string_optional_fields(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["objective"] = {"bad": "value"}
        outline["slides"][0]["notes"] = 123

        with self.assertRaisesRegex(ValueError, r"slide 1 objective must be a string"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_non_string_source_sections(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["source_sections"] = ["Section A", 2]

        with self.assertRaisesRegex(ValueError, r"slide 1 source_sections must be a list of strings"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_overlong_title_only_subtitle(self) -> None:
        outline = self.make_outline()
        outline["slides"][0]["layout"] = "title-only"
        outline["slides"][0]["objective"] = "这是一段明显会超过 title-only 副标题预算的超长文案" * 3

        with self.assertRaisesRegex(ValueError, r"title-only subtitle is too long"):
            doc_to_slides.validate_outline(outline)


class PublishRegressionTests(unittest.TestCase):
    def test_parse_args_rejects_removed_template_flag(self) -> None:
        with self.assertRaises(SystemExit):
            with redirect_stderr(io.StringIO()):
                doc_to_slides.parse_args(
                    [
                        "resolve-target",
                        "--run-dir",
                        "/tmp/run",
                        "--target-slides-url",
                        "https://example.feishu.cn/slides/sldcnTarget123",
                        "--template-slides-url",
                        "https://example.feishu.cn/wiki/template_token",
                    ]
                )

    def test_publish_new_deck_uses_default_create_output_shape(self) -> None:
        commands: list[list[str]] = []

        def fake_run_lark_cli(args: list[str]) -> dict:
            commands.append(args)
            return {
                "data": {
                    "xml_presentation_id": "sldcnCreated123",
                    "url": "https://example.feishu.cn/slides/sldcnCreated123",
                }
            }

        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.object(doc_to_slides, "run_lark_cli", side_effect=fake_run_lark_cli), patch.object(
                doc_to_slides,
                "create_slide_in_presentation",
                return_value="slide_001",
            ):
                result = doc_to_slides.publish_new_deck(
                    "Quarterly Review",
                    ["<slide><content /></slide>"],
                    Path(tmpdir),
                )

        self.assertEqual("new", result["target_mode"])
        self.assertEqual(
            [
                "lark-cli",
                "slides",
                "+create",
                "--as",
                "user",
                "--title",
                "Quarterly Review",
            ],
            commands[0],
        )

    def test_render_outline_infers_append_target_theme_from_resolved_target(self) -> None:
        outline = {
            "presentation": {
                "title": "Q2 Review",
                "target_mode": "append",
                "content_mode": "report",
                "source": {
                    "input_kind": "doc_url",
                    "resolved_kind": "doc_url",
                    "resolved_value": "https://example.feishu.cn/docs/doccn123",
                },
            },
            "slides": [
                {
                    "no": 1,
                    "role": "content",
                    "section_divider": False,
                    "title": "Status",
                    "layout": "title-body",
                    "key_points": ["Point A", "Point B"],
                }
            ],
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            doc_to_slides.write_json(
                run_dir / "resolved-target.json",
                {
                    "target_mode": "append",
                    "xml_presentation_id": "sldcnTarget123",
                    "url": "https://example.feishu.cn/slides/sldcnTarget123",
                },
            )

            def fake_run_lark_cli(args: list[str]) -> dict:
                if args[:4] == ["lark-cli", "slides", "xml_presentations", "get"]:
                    return {
                        "data": {
                            "xml_presentation": {
                                "content": (
                                    "<presentation xmlns=\"http://www.larkoffice.com/sml/2.0\">"
                                    "<slide><style><fill><fillColor color=\"rgb(15,23,42)\"/></fill></style></slide>"
                                    "</presentation>"
                                )
                            }
                        }
                    }
                raise AssertionError(f"unexpected command: {args}")

            with patch.object(doc_to_slides, "run_lark_cli", side_effect=fake_run_lark_cli):
                result = doc_to_slides.render_outline(outline, run_dir)
                summary = json.loads((run_dir / "render-summary.json").read_text(encoding="utf-8"))
                slides = result["slides"]

        self.assertEqual("spotlight", summary["theme"])
        self.assertEqual("spotlight", summary["append_inferred_theme"])
        self.assertIn('fillColor color="rgb(15,23,42)"', slides[0])

    def test_render_outline_infers_append_target_theme_from_rgba_background(self) -> None:
        outline = {
            "presentation": {
                "title": "Q2 Review",
                "target_mode": "append",
                "content_mode": "report",
                "source": {
                    "input_kind": "doc_url",
                    "resolved_kind": "doc_url",
                    "resolved_value": "https://example.feishu.cn/docs/doccn123",
                },
            },
            "slides": [
                {
                    "no": 1,
                    "role": "content",
                    "section_divider": False,
                    "title": "Status",
                    "layout": "title-body",
                    "key_points": ["Point A", "Point B"],
                }
            ],
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            doc_to_slides.write_json(
                run_dir / "resolved-target.json",
                {
                    "target_mode": "append",
                    "xml_presentation_id": "sldcnTarget123",
                    "url": "https://example.feishu.cn/slides/sldcnTarget123",
                },
            )

            def fake_run_lark_cli(args: list[str]) -> dict:
                if args[:4] == ["lark-cli", "slides", "xml_presentations", "get"]:
                    return {
                        "data": {
                            "xml_presentation": {
                                "content": (
                                    "<presentation xmlns=\"http://www.larkoffice.com/sml/2.0\">"
                                    "<slide><style><fill><fillColor color=\"rgba(15, 23, 42, 1)\"/></fill></style></slide>"
                                    "</presentation>"
                                )
                            }
                        }
                    }
                raise AssertionError(f"unexpected command: {args}")

            with patch.object(doc_to_slides, "run_lark_cli", side_effect=fake_run_lark_cli):
                result = doc_to_slides.render_outline(outline, run_dir)
                summary = json.loads((run_dir / "render-summary.json").read_text(encoding="utf-8"))
                slides = result["slides"]

        self.assertEqual("spotlight", summary["theme"])
        self.assertEqual("spotlight", summary["append_inferred_theme"])
        self.assertIn('fillColor color="rgb(15,23,42)"', slides[0])

    def test_render_outline_infers_append_target_theme_from_theme_background(self) -> None:
        outline = {
            "presentation": {
                "title": "Q2 Review",
                "target_mode": "append",
                "content_mode": "report",
                "source": {
                    "input_kind": "doc_url",
                    "resolved_kind": "doc_url",
                    "resolved_value": "https://example.feishu.cn/docs/doccn123",
                },
            },
            "slides": [
                {
                    "no": 1,
                    "role": "content",
                    "section_divider": False,
                    "title": "Status",
                    "layout": "title-body",
                    "key_points": ["Point A", "Point B"],
                }
            ],
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            doc_to_slides.write_json(
                run_dir / "resolved-target.json",
                {
                    "target_mode": "append",
                    "xml_presentation_id": "sldcnTarget123",
                    "url": "https://example.feishu.cn/slides/sldcnTarget123",
                },
            )

            def fake_run_lark_cli(args: list[str]) -> dict:
                if args[:4] == ["lark-cli", "slides", "xml_presentations", "get"]:
                    return {
                        "data": {
                            "xml_presentation": {
                                "content": (
                                    "<presentation xmlns=\"http://www.larkoffice.com/sml/2.0\">"
                                    "<theme><background><fillColor color=\"rgba(15, 23, 42, 1)\"/></background></theme>"
                                    "<slide></slide>"
                                    "</presentation>"
                                )
                            }
                        }
                    }
                raise AssertionError(f"unexpected command: {args}")

            with patch.object(doc_to_slides, "run_lark_cli", side_effect=fake_run_lark_cli):
                result = doc_to_slides.render_outline(outline, run_dir)
                summary = json.loads((run_dir / "render-summary.json").read_text(encoding="utf-8"))
                slides = result["slides"]

        self.assertEqual("spotlight", summary["theme"])
        self.assertEqual("spotlight", summary["append_inferred_theme"])
        self.assertIn('fillColor color="rgb(15,23,42)"', slides[0])

    def test_render_outline_records_layout_guard_summary(self) -> None:
        outline = {
            "presentation": {
                "title": "Q2 Review",
                "target_mode": "new",
                "content_mode": "report",
                "source": {
                    "input_kind": "doc_url",
                    "resolved_kind": "doc_url",
                    "resolved_value": "https://example.feishu.cn/docs/doccn123",
                },
            },
            "slides": [
                {
                    "no": 1,
                    "role": "content",
                    "section_divider": False,
                    "title": "Status",
                    "layout": "title-body",
                    "key_points": ["Point A", "Point B"],
                }
            ],
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            result = doc_to_slides.render_outline(outline, run_dir)
            summary = json.loads((run_dir / "render-summary.json").read_text(encoding="utf-8"))

        self.assertEqual(1, result["count"])
        self.assertEqual(1, summary["layout_guard"]["slides_checked"])
        self.assertEqual(960, summary["layout_guard"]["slide_size"]["width"])
        self.assertEqual(540, summary["layout_guard"]["slide_size"]["height"])

    def test_validate_rendered_slides_rejects_out_of_bounds_shape(self) -> None:
        slide_xml = (
            '<slide xmlns="http://www.larkoffice.com/sml/2.0">'
            '<style><fill><fillColor color="rgb(243,247,252)"/></fill></style>'
            '<data>'
            '<shape type="text" topLeftX="900" topLeftY="100" width="100" height="80">'
            '<content textType="title"><p>Overflow</p></content>'
            '</shape>'
            '</data>'
            '</slide>'
        )

        with self.assertRaisesRegex(ValueError, r"render layout guard failed"):
            doc_to_slides.validate_rendered_slides([slide_xml], ["title-body"])

    def test_validate_rendered_slides_rejects_title_body_without_body_text(self) -> None:
        slide_xml = (
            '<slide xmlns="http://www.larkoffice.com/sml/2.0">'
            '<style><fill><fillColor color="rgb(243,247,252)"/></fill></style>'
            '<data>'
            '<shape type="text" topLeftX="80" topLeftY="94" width="800" height="80">'
            '<content textType="title"><p>Status</p></content>'
            '</shape>'
            '</data>'
            '</slide>'
        )

        with self.assertRaisesRegex(ValueError, r"title-body layout is missing a body text shape"):
            doc_to_slides.validate_rendered_slides([slide_xml], ["title-body"])

    def test_validate_rendered_slides_rejects_two_column_without_second_panel(self) -> None:
        slide_xml = (
            '<slide xmlns="http://www.larkoffice.com/sml/2.0">'
            '<style><fill><fillColor color="rgb(243,247,252)"/></fill></style>'
            '<data>'
            '<shape type="text" topLeftX="80" topLeftY="94" width="800" height="80">'
            '<content textType="title"><p>Status</p></content>'
            '</shape>'
            '<shape type="rect" topLeftX="80" topLeftY="180" width="360" height="252"></shape>'
            '</data>'
            '</slide>'
        )

        with self.assertRaisesRegex(ValueError, r"two-column layout is missing its two content panels"):
            doc_to_slides.validate_rendered_slides([slide_xml], ["two-column"])

    def test_validate_rendered_slides_emits_sparse_density_warning(self) -> None:
        slide_xml = (
            '<slide xmlns="http://www.larkoffice.com/sml/2.0">'
            '<style><fill><fillColor color="rgb(243,247,252)"/></fill></style>'
            '<data>'
            '<shape type="text" topLeftX="80" topLeftY="40" width="320" height="60">'
            '<content textType="title"><p>Status</p></content>'
            '</shape>'
            '<shape type="text" topLeftX="80" topLeftY="120" width="320" height="60">'
            '<content textType="body"><p>Point A</p></content>'
            '</shape>'
            '</data>'
            '</slide>'
        )

        result = doc_to_slides.validate_rendered_slides([slide_xml], ["title-body"])

        self.assertTrue(any(warning["kind"] == "density_sparse" for warning in result["warnings"]))
        self.assertTrue(any(warning["severity"] == "low" for warning in result["warnings"]))
        self.assertTrue(any("visually sparse" in warning["message"] for warning in result["warnings"]))
        self.assertTrue(any(warning["kind"] == "density_sparse" for warning in result["per_slide"][0]["warnings"]))

    def test_validate_rendered_slides_emits_dense_density_warning(self) -> None:
        slide_xml = (
            '<slide xmlns="http://www.larkoffice.com/sml/2.0">'
            '<style><fill><fillColor color="rgb(243,247,252)"/></fill></style>'
            '<data>'
            '<shape type="text" topLeftX="40" topLeftY="30" width="880" height="80">'
            '<content textType="title"><p>Status</p></content>'
            '</shape>'
            '<shape type="text" topLeftX="40" topLeftY="120" width="880" height="390">'
            '<content textType="body"><p>Point A</p></content>'
            '</shape>'
            '</data>'
            '</slide>'
        )

        result = doc_to_slides.validate_rendered_slides([slide_xml], ["title-body"])

        self.assertTrue(any(warning["kind"] == "density_dense" for warning in result["warnings"]))
        self.assertTrue(any(warning["severity"] == "medium" for warning in result["warnings"]))
        self.assertTrue(any("visually dense" in warning["message"] for warning in result["warnings"]))
        self.assertTrue(any(warning["kind"] == "overflow_risk" for warning in result["warnings"]))


class FetchPaginationRegressionTests(unittest.TestCase):
    def test_fetch_source_rejects_non_advancing_next_offset(self) -> None:
        resolved = {
            "resolved_kind": "doc_url",
            "resolved_value": "https://example.feishu.cn/docx/doccn123",
            "title": "",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)

            def fake_run_lark_cli(args: list[str]) -> dict:
                return {
                    "data": {
                        "title": "Quarterly Review",
                        "markdown": "# Quarterly Review",
                        "has_more": True,
                        "next_offset": 0,
                    }
                }

            with patch.object(doc_to_slides, "run_lark_cli", side_effect=fake_run_lark_cli):
                with self.assertRaisesRegex(RuntimeError, r"fetch pagination did not advance"):
                    doc_to_slides.fetch_source(resolved, run_dir)


if __name__ == "__main__":
    unittest.main()
