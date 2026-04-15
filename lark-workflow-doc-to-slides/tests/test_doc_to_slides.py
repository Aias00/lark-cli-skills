from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch


MODULE_PATH = (
    Path(__file__).resolve().parents[1] / "scripts" / "doc_to_slides.py"
)
SPEC = importlib.util.spec_from_file_location("doc_to_slides", MODULE_PATH)
doc_to_slides = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(doc_to_slides)


def write_render_inputs(run_dir: Path, outline: dict, slides: list[str]) -> None:
    (run_dir / "render-summary.json").write_text(
        json.dumps(
            {
                "count": len(slides),
                "outline_fingerprint": doc_to_slides.fingerprint_payload(outline),
                "slides_fingerprint": doc_to_slides.fingerprint_payload(slides),
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )


def relative_luminance(color: str) -> float:
    def channel(component: int) -> float:
        normalized = component / 255
        if normalized <= 0.03928:
            return normalized / 12.92
        return ((normalized + 0.055) / 1.055) ** 2.4

    red, green, blue = doc_to_slides.parse_rgb_color(color)
    return 0.2126 * channel(red) + 0.7152 * channel(green) + 0.0722 * channel(blue)


def contrast_ratio(foreground: str, background: str) -> float:
    lighter = max(relative_luminance(foreground), relative_luminance(background))
    darker = min(relative_luminance(foreground), relative_luminance(background))
    return (lighter + 0.05) / (darker + 0.05)


class WikiResolutionTests(unittest.TestCase):
    def test_resolve_target_requires_target_url(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaisesRegex(
                ValueError,
                r"resolve-target requires target_slides_url",
            ):
                doc_to_slides.resolve_target(
                    None,
                    Path(tmpdir),
                )

    def test_resolve_source_normalizes_raw_wiki_token(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            resolved = doc_to_slides.resolve_source(
                doc_to_slides.argparse.Namespace(
                    doc_url=None,
                    doc_token="wikcnKQ1k3ptesttoken",
                    doc_name=None,
                    run_dir=tmpdir,
                ),
                Path(tmpdir),
            )

        self.assertEqual("wiki_token", resolved["resolved_kind"])
        self.assertEqual("wikcnKQ1k3ptesttoken", resolved["resolved_value"])

    def test_fetch_source_uses_resolved_wiki_docx_token(self) -> None:
        resolved = {
            "resolved_kind": "wiki_token",
            "resolved_value": "wikcnKQ1k3ptesttoken",
            "title": "",
        }
        calls: list[list[str]] = []

        def fake_run_lark_cli(args: list[str]) -> dict:
            calls.append(args)
            if args[:4] == ["lark-cli", "wiki", "spaces", "get_node"]:
                return {
                    "node": {
                        "obj_type": "docx",
                        "obj_token": "doxcnResolved123",
                        "title": "Quarterly Review",
                    }
                }
            if args[:3] == ["lark-cli", "docs", "+fetch"]:
                return {
                    "data": {
                        "title": "Quarterly Review",
                        "markdown": "# Quarterly Review",
                        "has_more": False,
                    }
                }
            raise AssertionError(f"unexpected command: {args}")

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            with patch.object(doc_to_slides, "run_lark_cli", side_effect=fake_run_lark_cli):
                result = doc_to_slides.fetch_source(resolved, run_dir)

            source_md = (run_dir / "source.md").read_text(encoding="utf-8")

        fetch_calls = [args for args in calls if args[:3] == ["lark-cli", "docs", "+fetch"]]
        self.assertEqual(1, len(fetch_calls))
        self.assertIn("doxcnResolved123", fetch_calls[0])
        self.assertEqual("doxcnResolved123", result["resolved_fetch_target"])
        self.assertEqual("Quarterly Review", result["wiki_node"]["title"])
        self.assertEqual("# Quarterly Review\n", source_md)

    def test_extract_wiki_node_accepts_data_node_envelope(self) -> None:
        payload = {
            "data": {
                "node": {
                    "obj_type": "docx",
                    "obj_token": "doccn123",
                    "title": "Quarterly Review",
                }
            }
        }

        node = doc_to_slides.extract_wiki_node(payload)

        self.assertEqual("docx", node["obj_type"])
        self.assertEqual("doccn123", node["obj_token"])

    def test_resolve_fetch_target_surfaces_wiki_scope_guidance(self) -> None:
        resolved = {
            "resolved_kind": "wiki_url",
            "resolved_value": "https://example.feishu.cn/wiki/wiki_token",
        }

        with patch.object(
            doc_to_slides,
            "run_lark_cli",
            side_effect=RuntimeError("permission denied"),
        ):
            with self.assertRaisesRegex(
                RuntimeError,
                r"wiki source.*wiki:wiki:readonly",
            ):
                doc_to_slides.resolve_fetch_target(resolved)

    def test_resolve_target_slides_url_explains_non_slides_wiki_target(self) -> None:
        with patch.object(
            doc_to_slides,
            "run_lark_cli",
            return_value={
                "node": {
                    "obj_type": "docx",
                    "obj_token": "doccn123",
                    "title": "Project Notes",
                }
            },
        ):
            with self.assertRaisesRegex(
                RuntimeError,
                r"Project Notes.*docx.*not slides",
                ):
                    doc_to_slides.resolve_target_slides_url(
                        "https://example.feishu.cn/wiki/wiki_token"
                    )

    def test_resolve_target_slides_url_returns_resolved_wiki_obj_token(self) -> None:
        with patch.object(
            doc_to_slides,
            "run_lark_cli",
            return_value={
                "node": {
                    "obj_type": "slides",
                    "obj_token": "sldcnResolved123",
                    "title": "Q2 Deck",
                }
            },
        ):
            presentation_id, url = doc_to_slides.resolve_target_slides_url(
                "https://example.feishu.cn/wiki/wiki_token"
            )

        self.assertEqual("sldcnResolved123", presentation_id)
        self.assertEqual("https://example.feishu.cn/wiki/wiki_token", url)

    def test_resolve_target_writes_preflight_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            with patch.object(
                doc_to_slides,
                "run_lark_cli",
                return_value={
                    "node": {
                        "obj_type": "slides",
                        "obj_token": "sldcn123",
                        "title": "Q2 Deck",
                    }
                },
            ):
                resolved = doc_to_slides.resolve_target(
                    "https://example.feishu.cn/wiki/wiki_token",
                    run_dir,
                )

            artifact = json.loads(
                (run_dir / "resolved-target.json").read_text(encoding="utf-8")
            )

        self.assertEqual("append", resolved["target_mode"])
        self.assertEqual("sldcn123", resolved["xml_presentation_id"])
        self.assertEqual(resolved, artifact)

class PublishAppendPreflightTests(unittest.TestCase):
    def setUp(self) -> None:
        self.outline = {
            "presentation": {
                "title": "Q2 Review",
                "target_mode": "append",
                "content_mode": "faithful",
                "source": {
                    "input_kind": "doc_url",
                    "resolved_kind": "doc_url",
                    "resolved_value": "https://example.feishu.cn/docs/doccn123",
                },
            },
            "slides": [
                {
                    "no": 1,
                    "role": "section",
                    "section_divider": True,
                    "title": "Highlights",
                    "layout": "title-only",
                    "key_points": [],
                }
            ],
        }
        self.slides = ["<slide><content /></slide>"]

    def test_publish_append_uses_resolved_target_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            write_render_inputs(run_dir, self.outline, self.slides)
            doc_to_slides.write_json(
                run_dir / "resolved-target.json",
                {
                    "target_mode": "append",
                    "xml_presentation_id": "sldcnArtifact123",
                    "url": "https://example.feishu.cn/wiki/wiki_token",
                },
            )

            with patch.object(
                doc_to_slides,
                "resolve_target_slides_url",
                side_effect=AssertionError("publish should not re-resolve append target"),
            ), patch.object(
                doc_to_slides,
                "create_slide_in_presentation",
                return_value="slide_001",
            ) as create_slide:
                result = doc_to_slides.publish_slides(
                    self.outline,
                    self.slides,
                    run_dir,
                    None,
                )

        create_slide.assert_called_once_with("sldcnArtifact123", self.slides[0])
        self.assertEqual("append", result["target_mode"])
        self.assertEqual("sldcnArtifact123", result["xml_presentation_id"])
        self.assertEqual("https://example.feishu.cn/wiki/wiki_token", result["url"])
        self.assertEqual(["slide_001"], result["slide_ids"])

    def test_publish_append_requires_resolved_target_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            write_render_inputs(run_dir, self.outline, self.slides)

            with patch.object(doc_to_slides, "create_slide_in_presentation") as create_slide:
                with self.assertRaisesRegex(
                    RuntimeError,
                    r"resolved-target\.json is required before append publish",
                ):
                    doc_to_slides.publish_slides(
                        self.outline,
                        self.slides,
                        run_dir,
                        None,
                    )

        create_slide.assert_not_called()

    def test_publish_append_rejects_target_url_mismatch_with_artifact(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            write_render_inputs(run_dir, self.outline, self.slides)
            doc_to_slides.write_json(
                run_dir / "resolved-target.json",
                {
                    "target_mode": "append",
                    "xml_presentation_id": "sldcnArtifact123",
                    "url": "https://example.feishu.cn/wiki/wiki_token",
                },
            )

            with patch.object(
                doc_to_slides,
                "resolve_target_slides_url",
                return_value=("sldcnOther999", "https://example.feishu.cn/slides/sldcnOther999"),
            ), patch.object(doc_to_slides, "create_slide_in_presentation") as create_slide:
                with self.assertRaisesRegex(
                    RuntimeError,
                    r"target_slides_url does not match resolved-target\.json",
                ):
                    doc_to_slides.publish_slides(
                        self.outline,
                        self.slides,
                        run_dir,
                        "https://example.feishu.cn/wiki/other_target",
                    )

        create_slide.assert_not_called()

    def test_publish_append_allows_equivalent_target_url_when_resolved_id_matches(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            write_render_inputs(run_dir, self.outline, self.slides)
            doc_to_slides.write_json(
                run_dir / "resolved-target.json",
                {
                    "target_mode": "append",
                    "xml_presentation_id": "sldcnArtifact123",
                    "url": "https://example.feishu.cn/wiki/wiki_token",
                },
            )

            with patch.object(
                doc_to_slides,
                "resolve_target_slides_url",
                return_value=("sldcnArtifact123", "https://example.feishu.cn/slides/sldcnArtifact123"),
            ), patch.object(
                doc_to_slides,
                "create_slide_in_presentation",
                return_value="slide_001",
            ) as create_slide:
                result = doc_to_slides.publish_slides(
                    self.outline,
                    self.slides,
                    run_dir,
                    "https://example.feishu.cn/slides/sldcnArtifact123",
                )

        create_slide.assert_called_once_with("sldcnArtifact123", self.slides[0])
        self.assertEqual(["slide_001"], result["slide_ids"])

    def test_publish_new_rejects_append_reference_inputs(self) -> None:
        outline = {
            "presentation": {
                "title": "Standalone Deck",
                "target_mode": "new",
                "content_mode": "faithful",
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
                    "key_points": ["Point A"],
                }
            ],
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            run_dir = Path(tmpdir)
            write_render_inputs(run_dir, outline, self.slides)

            with patch.object(doc_to_slides, "publish_new_deck") as publish_new:
                with self.assertRaisesRegex(
                    ValueError,
                    r"publish with target_mode=new does not accept target_slides_url",
                ):
                    doc_to_slides.publish_slides(
                        outline,
                        self.slides,
                        run_dir,
                        "https://example.feishu.cn/slides/sldcnTarget123",
                    )

        publish_new.assert_not_called()


class ThemePresetRenderingTests(unittest.TestCase):
    def make_outline(
        self,
        *,
        content_mode: str = "report",
        theme: str | None = None,
        subtitle: str | None = None,
        slides: list[dict] | None = None,
    ) -> dict:
        presentation = {
            "title": "Theme Review",
            "target_mode": "new",
            "content_mode": content_mode,
            "source": {
                "input_kind": "doc_url",
                "resolved_kind": "doc_url",
                "resolved_value": "https://example.feishu.cn/docs/doccn123",
            },
        }
        if theme is not None:
            presentation["theme"] = theme
        if subtitle is not None:
            presentation["subtitle"] = subtitle
        return {
            "presentation": presentation,
            "slides": slides
            or [
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

    def render_outline(self, outline: dict, *, run_dir: Path | None = None) -> tuple[dict, list[str]]:
        if run_dir is None:
            with tempfile.TemporaryDirectory() as tmpdir:
                local_run_dir = Path(tmpdir)
                result = doc_to_slides.render_outline(outline, local_run_dir)
                summary = json.loads((local_run_dir / "render-summary.json").read_text(encoding="utf-8"))
            return summary, result["slides"]

        result = doc_to_slides.render_outline(outline, run_dir)
        summary = json.loads((run_dir / "render-summary.json").read_text(encoding="utf-8"))
        return summary, result["slides"]

    def test_validate_outline_allows_builtin_theme(self) -> None:
        outline = self.make_outline(theme="spotlight")

        doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_unknown_theme(self) -> None:
        outline = self.make_outline(theme="unknown")

        with self.assertRaisesRegex(ValueError, r"invalid presentation.theme: unknown"):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_allows_builtin_cover_styles(self) -> None:
        for cover_style in ("editorial", "modern"):
            with self.subTest(cover_style=cover_style):
                outline = self.make_outline()
                outline["presentation"]["cover_style"] = cover_style

                doc_to_slides.validate_outline(outline)

    def test_validate_outline_rejects_unknown_cover_style(self) -> None:
        outline = self.make_outline()
        outline["presentation"]["cover_style"] = "retro"

        with self.assertRaisesRegex(ValueError, r"invalid presentation.cover_style: retro"):
            doc_to_slides.validate_outline(outline)

    def test_render_outline_defaults_theme_from_report_content_mode(self) -> None:
        summary, slides = self.render_outline(self.make_outline(content_mode="report"))

        self.assertEqual("briefing", summary["theme"])
        self.assertIn('fillColor color="rgb(243,247,252)"', slides[0])

    def test_render_outline_defaults_theme_from_faithful_content_mode(self) -> None:
        summary, slides = self.render_outline(self.make_outline(content_mode="faithful"))

        self.assertEqual("document", summary["theme"])
        self.assertIn('fillColor color="rgb(250,247,240)"', slides[0])

    def test_render_outline_explicit_theme_overrides_content_mode_default(self) -> None:
        summary, slides = self.render_outline(
            self.make_outline(content_mode="faithful", theme="spotlight")
        )

        self.assertEqual("spotlight", summary["theme"])
        self.assertIn('fillColor color="rgb(15,23,42)"', slides[0])
        self.assertNotIn('fillColor color="rgb(250,247,240)"', slides[0])

    def test_render_outline_defaults_cover_style_from_theme(self) -> None:
        report_summary, _ = self.render_outline(self.make_outline(content_mode="report"))
        faithful_summary, _ = self.render_outline(self.make_outline(content_mode="faithful"))
        spotlight_summary, _ = self.render_outline(self.make_outline(theme="spotlight"))

        self.assertEqual("editorial", report_summary["cover_style"])
        self.assertEqual("editorial", faithful_summary["cover_style"])
        self.assertEqual("modern", spotlight_summary["cover_style"])

    def test_render_outline_explicit_cover_style_overrides_theme_default(self) -> None:
        outline = self.make_outline(
            theme="spotlight",
            slides=[
                {
                    "no": 1,
                    "role": "cover",
                    "section_divider": False,
                    "title": "Quarterly Review",
                    "layout": "title-only",
                    "key_points": ["Q2 Results"],
                    "objective": "Executive briefing",
                }
            ],
        )
        default_summary, default_slides = self.render_outline(self.make_outline(theme="spotlight", slides=outline["slides"]))
        outline["presentation"]["cover_style"] = "editorial"
        summary, slides = self.render_outline(outline)

        self.assertEqual("modern", default_summary["cover_style"])
        self.assertEqual("editorial", summary["cover_style"])
        self.assertNotEqual(default_slides[0], slides[0])
        self.assertIn('fillColor color="rgb(249,115,22)"', slides[0])

    def test_render_outline_only_changes_cover_rendering_across_cover_styles(self) -> None:
        slides = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly Review",
                "layout": "title-only",
                "key_points": ["Q2 Results"],
                "objective": "Executive briefing",
            },
            {
                "no": 2,
                "role": "content",
                "section_divider": False,
                "title": "Status",
                "layout": "title-body",
                "key_points": ["Point A", "Point B"],
            },
        ]
        editorial_outline = self.make_outline(theme="spotlight", slides=slides)
        editorial_outline["presentation"]["cover_style"] = "editorial"
        modern_outline = self.make_outline(theme="spotlight", slides=slides)
        modern_outline["presentation"]["cover_style"] = "modern"

        editorial_summary, editorial_slides = self.render_outline(editorial_outline)
        modern_summary, modern_slides = self.render_outline(modern_outline)

        self.assertEqual("editorial", editorial_summary["cover_style"])
        self.assertEqual("modern", modern_summary["cover_style"])
        self.assertNotEqual(editorial_summary["slides_fingerprint"], modern_summary["slides_fingerprint"])
        self.assertNotEqual(editorial_slides[0], modern_slides[0])
        self.assertEqual(editorial_slides[1], modern_slides[1])

    def test_render_outline_changes_fingerprint_across_themes(self) -> None:
        report_summary, report_slides = self.render_outline(
            self.make_outline(content_mode="report")
        )
        faithful_summary, faithful_slides = self.render_outline(
            self.make_outline(content_mode="faithful")
        )

        self.assertNotEqual(report_summary["slides_fingerprint"], faithful_summary["slides_fingerprint"])
        self.assertNotEqual(report_slides[0], faithful_slides[0])

    def test_get_cover_subtitle_prefers_presentation_subtitle_then_objective_then_key_point(self) -> None:
        slide = {
            "title": "Quarterly Review",
            "objective": "Slide objective",
            "key_points": ["Key point subtitle"],
        }

        self.assertEqual(
            "Presentation subtitle",
            doc_to_slides.get_cover_subtitle({"subtitle": "Presentation subtitle"}, slide),
        )
        self.assertEqual(
            "Slide objective",
            doc_to_slides.get_cover_subtitle({}, slide),
        )
        self.assertEqual(
            "Key point subtitle",
            doc_to_slides.get_cover_subtitle({}, {"title": "Quarterly Review", "key_points": ["Key point subtitle"]}),
        )

    def test_render_cover_uses_presentation_subtitle_source(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly Review",
                "layout": "title-only",
                "key_points": ["Fallback key point"],
                "objective": "Fallback objective",
            }
        ]
        outline = self.make_outline(
            content_mode="report",
            subtitle="Presentation subtitle wins",
            slides=cover_slide,
        )
        outline["presentation"]["cover_style"] = "modern"

        _, slides = self.render_outline(outline)

        self.assertIn("Presentation subtitle wins", slides[0])
        self.assertNotIn("Fallback objective", slides[0])
        self.assertNotIn("Fallback key point", slides[0])

    def test_render_editorial_cover_uses_compact_rule_without_legacy_banner(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly Review",
                "layout": "title-only",
                "key_points": ["Q2 Results"],
                "objective": "Executive briefing",
            }
        ]

        _, briefing_slides = self.render_outline(
            self.make_outline(content_mode="report", slides=cover_slide)
        )

        self.assertIn('topLeftX="96" topLeftY="96" width="188" height="10"', briefing_slides[0])
        self.assertNotIn('topLeftX="80" topLeftY="82" width="800" height="124"', briefing_slides[0])
        self.assertIn('color="rgb(15,23,42)"', briefing_slides[0])
        self.assertIn('color="rgb(51,65,85)"', briefing_slides[0])
        self.assertNotIn('fillColor color="rgb(255,255,255)"', briefing_slides[0])

    def test_render_modern_cover_uses_briefing_hero_and_card_surfaces(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly Review",
                "layout": "title-only",
                "key_points": ["Q2 Results"],
                "objective": "Executive briefing",
            }
        ]
        outline = self.make_outline(content_mode="report", slides=cover_slide)
        outline["presentation"]["cover_style"] = "modern"

        _, slides = self.render_outline(outline)

        theme = doc_to_slides.THEME_PRESETS["briefing"]
        self.assertIn('topLeftX="80" topLeftY="88" width="18" height="288"', slides[0])
        self.assertIn(
            f'topLeftX="120" topLeftY="96" width="628" height="182"><fill><fillColor color="{theme["cover_modern_hero_fill"]}"',
            slides[0],
        )
        self.assertIn(
            f'topLeftX="144" topLeftY="290" width="548" height="84"><fill><fillColor color="{theme["cover_modern_card_fill"]}"',
            slides[0],
        )
        self.assertIn(f'<span color="{theme["title_color"]}"', slides[0])
        self.assertIn(f'<span color="{theme["body_color"]}"', slides[0])
        self.assertNotIn(f'<span color="{theme["cover_title_color"]}"', slides[0])
        self.assertNotIn(f'<span color="{theme["cover_body_color"]}"', slides[0])
        self.assertGreaterEqual(contrast_ratio(theme["title_color"], theme["cover_modern_hero_fill"]), 4.5)
        self.assertGreaterEqual(contrast_ratio(theme["body_color"], theme["cover_modern_card_fill"]), 4.5)

    def test_render_modern_cover_uses_document_hero_and_card_surfaces(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly Review",
                "layout": "title-only",
                "key_points": ["Q2 Results"],
                "objective": "Executive briefing",
            }
        ]
        outline = self.make_outline(content_mode="faithful", slides=cover_slide)
        outline["presentation"]["cover_style"] = "modern"

        _, slides = self.render_outline(outline)

        theme = doc_to_slides.THEME_PRESETS["document"]
        self.assertIn(
            f'topLeftX="120" topLeftY="96" width="628" height="182"><fill><fillColor color="{theme["cover_modern_hero_fill"]}"',
            slides[0],
        )
        self.assertIn(
            f'topLeftX="144" topLeftY="290" width="548" height="84"><fill><fillColor color="{theme["cover_modern_card_fill"]}"',
            slides[0],
        )
        self.assertIn(f'<span color="{theme["title_color"]}"', slides[0])
        self.assertIn(f'<span color="{theme["body_color"]}"', slides[0])
        self.assertGreaterEqual(contrast_ratio(theme["title_color"], theme["cover_modern_hero_fill"]), 4.5)
        self.assertGreaterEqual(contrast_ratio(theme["body_color"], theme["cover_modern_card_fill"]), 4.5)

    def test_render_modern_cover_keeps_inverse_pairs_on_spotlight_theme(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly Review",
                "layout": "title-only",
                "key_points": ["Q2 Results"],
                "objective": "Executive briefing",
            }
        ]

        _, slides = self.render_outline(
            self.make_outline(theme="spotlight", slides=cover_slide)
        )

        theme = doc_to_slides.THEME_PRESETS["spotlight"]
        self.assertIn(
            f'topLeftX="120" topLeftY="96" width="628" height="182"><fill><fillColor color="{theme["cover_modern_hero_fill"]}"',
            slides[0],
        )
        self.assertIn(
            f'topLeftX="144" topLeftY="290" width="548" height="84"><fill><fillColor color="{theme["cover_modern_card_fill"]}"',
            slides[0],
        )
        self.assertIn(f'<span color="{theme["cover_title_color"]}"', slides[0])
        self.assertIn(f'<span color="{theme["cover_body_color"]}"', slides[0])
        self.assertGreaterEqual(contrast_ratio(theme["cover_title_color"], theme["cover_modern_hero_fill"]), 4.5)
        self.assertGreaterEqual(contrast_ratio(theme["cover_body_color"], theme["cover_modern_card_fill"]), 3.0)

    def test_render_modern_cover_expands_for_long_briefing_cover_copy(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Quarterly business review with cross-functional operating changes, launch sequencing, and region-by-region readiness details",
                "layout": "title-only",
                "key_points": ["Fallback key point"],
            }
        ]
        outline = self.make_outline(
            content_mode="report",
            subtitle="A detailed executive subtitle covering dependencies, rollout sequencing, adoption constraints, and commercial readiness for the next planning cycle.",
            slides=cover_slide,
        )
        outline["presentation"]["cover_style"] = "modern"

        _, slides = self.render_outline(outline)

        self.assertIn('fontSize="24"', slides[0])
        self.assertIn('topLeftX="144" topLeftY="122" width="580" height="152"', slides[0])
        self.assertIn('topLeftX="144" topLeftY="290" width="548" height="116"', slides[0])
        self.assertIn('topLeftX="168" topLeftY="308" width="500" height="80"', slides[0])

    def test_render_modern_cover_expands_for_long_document_cover_copy(self) -> None:
        cover_slide = [
            {
                "no": 1,
                "role": "cover",
                "section_divider": False,
                "title": "Operational review and editorial synthesis for compliance, implementation sequencing, and governance milestones across the organization",
                "layout": "title-only",
                "key_points": ["Fallback key point"],
            }
        ]
        outline = self.make_outline(
            content_mode="faithful",
            subtitle="A longer subtitle for the document theme that needs extra height for line wrapping while preserving the modern hero-card structure.",
            slides=cover_slide,
        )
        outline["presentation"]["cover_style"] = "modern"

        _, slides = self.render_outline(outline)

        self.assertIn('fontSize="24"', slides[0])
        self.assertIn('topLeftX="144" topLeftY="122" width="580" height="152"', slides[0])
        self.assertIn('topLeftX="144" topLeftY="290" width="548" height="116"', slides[0])
        self.assertIn('topLeftX="168" topLeftY="308" width="500" height="80"', slides[0])

    def test_render_outline_uses_theme_tokens_across_layouts(self) -> None:
        outline = self.make_outline(
            theme="spotlight",
            slides=[
                {
                    "no": 1,
                    "role": "cover",
                    "section_divider": False,
                    "title": "Quarterly Review",
                    "layout": "title-only",
                    "key_points": ["Q2 Results"],
                    "objective": "Executive briefing",
                },
                {
                    "no": 2,
                    "role": "content",
                    "section_divider": False,
                    "title": "Highlights",
                    "layout": "title-body",
                    "key_points": ["Launch complete", "Risk retired"],
                },
                {
                    "no": 3,
                    "role": "comparison",
                    "section_divider": False,
                    "title": "Options",
                    "layout": "comparison",
                    "key_points": ["Lower cost", "Faster rollout", "Higher risk", "More control"],
                    "source_sections": ["Build", "Buy"],
                },
                {
                    "no": 4,
                    "role": "timeline",
                    "section_divider": False,
                    "title": "Plan",
                    "layout": "timeline",
                    "key_points": ["Week 1", "Week 2", "Week 3"],
                },
                {
                    "no": 5,
                    "role": "metrics",
                    "section_divider": False,
                    "title": "KPIs",
                    "layout": "metrics",
                    "key_points": ["99.9% SLA", "12d lead time", "31% adoption"],
                },
            ],
        )

        _, slides = self.render_outline(outline)

        self.assertIn('fillColor color="rgb(249,115,22)"', slides[0])
        self.assertIn('fillColor color="rgb(56,189,248)"', slides[1])
        self.assertIn('fillColor color="rgb(56,189,248)"', slides[2])
        self.assertIn('fillColor color="rgb(30,41,59)"', slides[2])
        self.assertIn('fillColor color="rgb(51,65,85)"', slides[2])
        self.assertIn('fillColor color="rgb(234,88,12)"', slides[3])
        self.assertIn('fillColor color="rgb(124,58,237)"', slides[4])

    def test_validate_outline_keeps_append_cover_guard_with_theme(self) -> None:
        outline = self.make_outline(
            content_mode="report",
            theme="briefing",
            slides=[
                {
                    "no": 1,
                    "role": "cover",
                    "section_divider": False,
                    "title": "Executive Summary",
                    "layout": "title-only",
                    "key_points": ["Summary"],
                }
            ],
        )
        outline["presentation"]["target_mode"] = "append"

        with self.assertRaisesRegex(
            ValueError,
            r"append mode cannot include cover slides",
        ):
            doc_to_slides.validate_outline(outline)

    def test_validate_outline_keeps_append_cover_guard_with_cover_style(self) -> None:
        outline = self.make_outline(
            content_mode="report",
            theme="briefing",
            slides=[
                {
                    "no": 1,
                    "role": "cover",
                    "section_divider": False,
                    "title": "Executive Summary",
                    "layout": "title-only",
                    "key_points": ["Summary"],
                }
            ],
        )
        outline["presentation"]["target_mode"] = "append"
        outline["presentation"]["cover_style"] = "modern"

        with self.assertRaisesRegex(
            ValueError,
            r"append mode cannot include cover slides",
        ):
            doc_to_slides.validate_outline(outline)

    def test_render_section_divider_takes_precedence_over_cover_styling(self) -> None:
        outline = self.make_outline(
            content_mode="report",
            slides=[
                {
                    "no": 1,
                    "role": "section",
                    "section_divider": True,
                    "title": "Highlights",
                    "layout": "title-only",
                    "key_points": ["Section opener"],
                }
            ],
        )
        outline["presentation"]["target_mode"] = "append"

        _, slides = self.render_outline(outline)

        self.assertIn('fillColor color="rgb(59,130,246)"', slides[0])
        self.assertNotIn('fillColor color="rgb(37,99,235)"', slides[0])

    def test_render_section_divider_takes_precedence_over_cover_style(self) -> None:
        outline = self.make_outline(
            theme="spotlight",
            slides=[
                {
                    "no": 1,
                    "role": "section",
                    "section_divider": True,
                    "title": "Highlights",
                    "layout": "title-only",
                    "key_points": ["Section opener"],
                }
            ],
        )
        outline["presentation"]["target_mode"] = "append"
        outline["presentation"]["cover_style"] = "modern"

        _, slides = self.render_outline(outline)

        self.assertIn('fillColor color="rgb(56,189,248)"', slides[0])
        self.assertNotIn('fillColor color="rgb(249,115,22)"', slides[0])

    def test_render_non_cover_title_only_and_section_divider_are_unchanged_across_cover_styles(self) -> None:
        slides = [
            {
                "no": 1,
                "role": "content",
                "section_divider": False,
                "title": "Status",
                "layout": "title-only",
                "key_points": ["Stable rollout"],
            },
            {
                "no": 2,
                "role": "section",
                "section_divider": True,
                "title": "Highlights",
                "layout": "title-only",
                "key_points": ["Section opener"],
            },
        ]
        editorial_outline = self.make_outline(theme="spotlight", slides=slides)
        editorial_outline["presentation"]["cover_style"] = "editorial"
        modern_outline = self.make_outline(theme="spotlight", slides=slides)
        modern_outline["presentation"]["cover_style"] = "modern"

        _, editorial_slides = self.render_outline(editorial_outline)
        _, modern_slides = self.render_outline(modern_outline)

        self.assertEqual(editorial_slides[0], modern_slides[0])
        self.assertEqual(editorial_slides[1], modern_slides[1])


if __name__ == "__main__":
    unittest.main()
