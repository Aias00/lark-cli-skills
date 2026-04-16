#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
import unicodedata
from pathlib import Path
from urllib.parse import urlparse
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET


VALID_LAYOUTS = {
    "title-only",
    "title-body",
    "two-column",
    "bullets",
    "comparison",
    "timeline",
    "metrics",
}
TARGET_MODES = {"new", "append"}
CONTENT_MODES = {"faithful", "report"}
DEFAULT_THEME_BY_CONTENT_MODE = {
    "faithful": "document",
    "report": "briefing",
}
COVER_STYLES = {"editorial", "modern"}
DEFAULT_COVER_STYLE_BY_THEME = {
    "briefing": "editorial",
    "document": "editorial",
    "spotlight": "modern",
    "minimal": "modern",
    "sunset": "modern",
}
THEME_PRESETS = {
    "briefing": {
        # 商务汇报风：专业蓝色系，适合周报、月报、项目汇报
        "slide_bg": "rgb(248,250,252)",
        "title_color": "rgb(15,23,42)",
        "body_color": "rgb(51,65,85)",
        "cover_title_color": "rgb(255,255,255)",
        "cover_body_color": "rgb(191,219,254)",
        "cover_modern_hero_fill": "rgb(29,78,216)",
        "cover_modern_card_fill": "rgb(255,255,255)",
        "muted_color": "rgb(100,116,139)",
        "panel_fill": "rgb(255,255,255)",
        "panel_alt_fill": "rgb(239,246,255)",
        "border_color": "rgb(203,213,225)",
        "cover_band_fill": "rgb(29,78,216)",
        "content_band_fill": "rgb(59,130,246)",
        "comparison_left_fill": "rgb(219,234,254)",
        "comparison_right_fill": "rgb(191,219,254)",
        "timeline_rail_fill": "rgb(59,130,246)",
        "timeline_item_fill": "rgb(255,255,255)",
        "metrics_card_fill": "rgb(219,234,254)",
        "metrics_highlight_fill": "rgb(59,130,246)",
        # 新增字段
        "sidebar_accent": "rgb(29,78,216)",
        "number_badge_fill": "rgb(29,78,216)",
        "number_badge_text": "rgb(255,255,255)",
        "title_bg_fill": "rgb(239,246,255)",
    },
    "document": {
        # 现代科技风：深蓝主色 + 浅灰背景，适合技术文档、知识分享
        "slide_bg": "rgb(248,250,252)",
        "title_color": "rgb(30,41,59)",
        "body_color": "rgb(71,85,105)",
        "cover_title_color": "rgb(255,255,255)",
        "cover_body_color": "rgb(148,163,184)",
        "cover_modern_hero_fill": "rgb(30,64,175)",
        "cover_modern_card_fill": "rgb(255,255,255)",
        "muted_color": "rgb(100,116,139)",
        "panel_fill": "rgb(255,255,255)",
        "panel_alt_fill": "rgb(241,245,249)",
        "border_color": "rgb(203,213,225)",
        "cover_band_fill": "rgb(30,64,175)",
        "content_band_fill": "rgb(59,130,246)",
        "comparison_left_fill": "rgb(239,246,255)",
        "comparison_right_fill": "rgb(224,231,255)",
        "timeline_rail_fill": "rgb(59,130,246)",
        "timeline_item_fill": "rgb(255,255,255)",
        "metrics_card_fill": "rgb(239,246,255)",
        "metrics_highlight_fill": "rgb(96,165,250)",
        # 新增字段
        "sidebar_accent": "rgb(59,130,246)",
        "number_badge_fill": "rgb(59,130,246)",
        "number_badge_text": "rgb(255,255,255)",
        "title_bg_fill": "rgb(239,246,255)",
    },
    "spotlight": {
        # 深色演示风：深色背景 + 鲜艳强调色，适合发布会、演讲
        "slide_bg": "rgb(15,23,42)",
        "title_color": "rgb(248,250,252)",
        "body_color": "rgb(226,232,240)",
        "cover_title_color": "rgb(255,255,255)",
        "cover_body_color": "rgb(254,215,170)",
        "cover_modern_hero_fill": "rgb(88,28,135)",
        "cover_modern_card_fill": "rgb(30,41,59)",
        "muted_color": "rgb(148,163,184)",
        "panel_fill": "rgb(30,41,59)",
        "panel_alt_fill": "rgb(51,65,85)",
        "border_color": "rgb(71,85,105)",
        "cover_band_fill": "rgb(168,85,247)",
        "content_band_fill": "rgb(139,92,246)",
        "comparison_left_fill": "rgb(30,41,59)",
        "comparison_right_fill": "rgb(51,65,85)",
        "timeline_rail_fill": "rgb(168,85,247)",
        "timeline_item_fill": "rgb(30,41,59)",
        "metrics_card_fill": "rgb(51,65,85)",
        "metrics_highlight_fill": "rgb(168,85,247)",
        # 新增字段
        "sidebar_accent": "rgb(139,92,246)",
        "number_badge_fill": "rgb(139,92,246)",
        "number_badge_text": "rgb(255,255,255)",
        "title_bg_fill": "rgb(30,41,59)",
    },
    "minimal": {
        # 极简黑白灰：设计感、艺术、现代简约风格
        "slide_bg": "rgb(255,255,255)",
        "title_color": "rgb(17,24,39)",
        "body_color": "rgb(55,65,81)",
        "cover_title_color": "rgb(255,255,255)",
        "cover_body_color": "rgb(156,163,175)",
        "cover_modern_hero_fill": "rgb(17,24,39)",
        "cover_modern_card_fill": "rgb(249,250,251)",
        "muted_color": "rgb(107,114,128)",
        "panel_fill": "rgb(249,250,251)",
        "panel_alt_fill": "rgb(243,244,246)",
        "border_color": "rgb(209,213,219)",
        "cover_band_fill": "rgb(17,24,39)",
        "content_band_fill": "rgb(75,85,99)",
        "comparison_left_fill": "rgb(243,244,246)",
        "comparison_right_fill": "rgb(229,231,235)",
        "timeline_rail_fill": "rgb(75,85,99)",
        "timeline_item_fill": "rgb(255,255,255)",
        "metrics_card_fill": "rgb(243,244,246)",
        "metrics_highlight_fill": "rgb(75,85,99)",
        # 新增字段
        "sidebar_accent": "rgb(75,85,99)",
        "number_badge_fill": "rgb(55,65,81)",
        "number_badge_text": "rgb(255,255,255)",
        "title_bg_fill": "rgb(243,244,246)",
    },
    "sunset": {
        # 暖橙琥珀色系：温暖、活力、创意场景
        "slide_bg": "rgb(255,251,245)",
        "title_color": "rgb(41,37,36)",
        "body_color": "rgb(68,64,60)",
        "cover_title_color": "rgb(255,255,255)",
        "cover_body_color": "rgb(254,243,199)",
        "cover_modern_hero_fill": "rgb(234,88,12)",
        "cover_modern_card_fill": "rgb(255,255,255)",
        "muted_color": "rgb(120,113,108)",
        "panel_fill": "rgb(255,255,255)",
        "panel_alt_fill": "rgb(254,252,247)",
        "border_color": "rgb(253,230,138)",
        "cover_band_fill": "rgb(234,88,12)",
        "content_band_fill": "rgb(245,158,11)",
        "comparison_left_fill": "rgb(254,243,199)",
        "comparison_right_fill": "rgb(254,215,170)",
        "timeline_rail_fill": "rgb(245,158,11)",
        "timeline_item_fill": "rgb(255,255,255)",
        "metrics_card_fill": "rgb(254,243,199)",
        "metrics_highlight_fill": "rgb(245,158,11)",
        # 新增字段
        "sidebar_accent": "rgb(234,88,12)",
        "number_badge_fill": "rgb(234,88,12)",
        "number_badge_text": "rgb(255,255,255)",
        "title_bg_fill": "rgb(254,243,199)",
    },
}
FETCHABLE_ENTITY_TYPES = {"DOC", "DOCX"}
RESOLVABLE_ENTITY_TYPES = {"WIKI"}
FETCHABLE_WIKI_OBJECT_TYPES = {"doc", "docx"}
WIKI_RESOLVED_KINDS = {"wiki_url", "wiki_token"}
WIKI_READ_SCOPE = "wiki:wiki:readonly"
SML_NS = "http://www.larkoffice.com/sml/2.0"
SLIDE_WIDTH = 960
SLIDE_HEIGHT = 540
HTML_TAG_RE = re.compile(r"<[^>]+>")
RGB_COLOR_RE = re.compile(
    r"rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*(?:0|1|0?\.\d+))?\s*\)"
)
PRESENTATION_TITLE_UNIT_LIMIT = 80
PRESENTATION_SUBTITLE_UNIT_LIMIT = 180
SLIDE_TITLE_UNIT_LIMITS = {
    "title-only": 56,
    "title-body": 44,
    "bullets": 44,
    "two-column": 36,
    "comparison": 36,
    "timeline": 40,
    "metrics": 28,
}
KEY_POINT_UNIT_LIMITS = {
    "title-only": 70,
    "title-body": 58,
    "bullets": 58,
    "two-column": 34,
    "comparison": 30,
    "timeline": 48,
    "metrics": 18,
}
TOTAL_KEY_POINT_UNIT_LIMITS = {
    "title-only": 140,
    "title-body": 190,
    "bullets": 190,
    "two-column": 120,
    "comparison": 120,
    "timeline": 160,
    "metrics": 72,
}
TITLE_ONLY_SUBTITLE_UNIT_LIMIT = 72


def extract_xml_presentation_payload(raw: dict) -> dict:
    data = raw.get("data")
    if isinstance(data, dict):
        if isinstance(data.get("xml_presentation"), dict):
            return data["xml_presentation"]
        if isinstance(data.get("content"), str):
            return data
    if isinstance(raw.get("xml_presentation"), dict):
        return raw["xml_presentation"]
    if isinstance(raw.get("content"), str):
        return raw
    raise RuntimeError("unexpected xml_presentation.get response shape")


class PublishError(RuntimeError):
    def __init__(self, message: str, result: dict) -> None:
        super().__init__(message)
        self.result = result


def emit_error(command: str, error: Exception) -> int:
    payload = {
        "ok": False,
        "command": command,
        "error": str(error),
        "error_type": type(error).__name__,
    }
    json.dump(payload, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    return 1


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Resolve, fetch, validate, render, and publish doc-to-slides workflow artifacts."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    resolve = subparsers.add_parser("resolve-source")
    resolve.add_argument("--doc-url")
    resolve.add_argument("--doc-token")
    resolve.add_argument("--doc-name")
    resolve.add_argument("--run-dir", required=True)

    target = subparsers.add_parser("resolve-target")
    target.add_argument("--target-slides-url")
    target.add_argument("--run-dir", required=True)

    choose = subparsers.add_parser("choose-source")
    choose.add_argument("--resolved-source", required=True)
    choose.add_argument("--candidate-index", required=True, type=int)

    fetch = subparsers.add_parser("fetch")
    fetch.add_argument("--resolved-source", required=True)
    fetch.add_argument("--run-dir", required=True)

    validate = subparsers.add_parser("validate-outline")
    validate.add_argument("--outline", required=True)

    render = subparsers.add_parser("render")
    render.add_argument("--outline", required=True)
    render.add_argument("--run-dir", required=True)

    publish = subparsers.add_parser("publish")
    publish.add_argument("--outline", required=True)
    publish.add_argument("--slides-json", required=True)
    publish.add_argument("--run-dir", required=True)
    publish.add_argument("--target-slides-url")

    return parser.parse_args(argv)


def ensure_run_dir(path_arg: str | None) -> Path:
    if not path_arg:
        raise ValueError("run directory is required")
    run_dir = Path(path_arg)
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def read_json(path: Path) -> object:
    return json.loads(path.read_text(encoding="utf-8"))


def run_lark_cli(args: list[str]) -> dict:
    proc = subprocess.run(args, capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        message = proc.stderr.strip() or proc.stdout.strip() or f"lark-cli failed: {' '.join(args)}"
        raise RuntimeError(message)
    try:
        return json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"lark-cli returned invalid JSON: {exc}") from exc


def normalize_entity_type(value: object) -> str:
    if isinstance(value, list):
        for item in value:
            normalized = normalize_entity_type(item)
            if normalized:
                return normalized
        return ""
    if value is None:
        return ""
    return str(value).strip().upper()


def strip_markup(text: object) -> str:
    if text is None:
        return ""
    return HTML_TAG_RE.sub("", str(text))


def stable_json_dumps(payload: object) -> str:
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def fingerprint_payload(payload: object) -> str:
    return hashlib.sha256(stable_json_dumps(payload).encode("utf-8")).hexdigest()


def text_units(text: str) -> int:
    collapsed = " ".join(text.split())
    return sum(2 if unicodedata.east_asian_width(char) in {"W", "F"} else 1 for char in collapsed)


def validate_text_budget(label: str, text: str, max_units: int) -> None:
    units = text_units(text)
    if units > max_units:
        raise ValueError(f"{label} is too long for the current layout budget: {units} > {max_units}")


def effective_title_only_subtitle(slide: dict) -> str:
    if isinstance(slide.get("objective"), str) and slide["objective"]:
        return slide["objective"]
    key_points = slide.get("key_points") or []
    if key_points:
        first = key_points[0]
        if isinstance(first, str):
            return first
    return ""


def extract_search_results(search_result: dict) -> list[dict]:
    data = search_result.get("data")
    if isinstance(data, dict):
        return data.get("results") or data.get("res_units") or []
    return search_result.get("results") or search_result.get("res_units") or []


def extract_fetch_payload(raw: dict) -> dict:
    data = raw.get("data")
    if isinstance(data, dict) and (
        "markdown" in data or "title" in data or "has_more" in data or "next_offset" in data
    ):
        return data
    return raw


def extract_search_candidates(search_result: dict) -> list[dict]:
    candidates: list[dict] = []
    for item in extract_search_results(search_result):
        result_meta = item.get("result_meta") or {}
        entity_type = normalize_entity_type(item.get("entity_type") or result_meta.get("doc_types"))
        url = result_meta.get("url") or item.get("url")
        if not url:
            continue
        if entity_type in FETCHABLE_ENTITY_TYPES:
            resolved_kind = "doc_url"
        elif entity_type in RESOLVABLE_ENTITY_TYPES:
            resolved_kind = "wiki_url"
        else:
            continue
        candidates.append(
            {
                "title": strip_markup(item.get("title") or item.get("title_highlighted") or result_meta.get("title") or ""),
                "resolved_kind": resolved_kind,
                "resolved_value": url,
                "entity_type": entity_type,
            }
        )
    return candidates


def extract_token(value: str) -> str:
    parsed = urlparse(value)
    if parsed.scheme:
        return parsed.path.rstrip("/").split("/")[-1]
    return value.rstrip("/").split("/")[-1]


def is_wiki_reference(value: str) -> bool:
    return "/wiki/" in value


def is_wiki_token(value: str) -> bool:
    return bool(value) and not urlparse(value).scheme and extract_token(value).lower().startswith("wik")


def looks_like_wiki_node(payload: object) -> bool:
    if not isinstance(payload, dict):
        return False
    return any(key in payload for key in ("obj_type", "obj_token", "node_token", "space_id"))


def extract_wiki_node(payload: dict) -> dict:
    candidates: list[object] = [payload.get("node")]
    data = payload.get("data")
    if isinstance(data, dict):
        candidates.append(data.get("node"))
        candidates.append(data)
    for candidate in candidates:
        if looks_like_wiki_node(candidate):
            return candidate
    raise RuntimeError("unexpected wiki get_node response shape")


def resolve_wiki_node(reference: str, *, context: str) -> dict:
    wiki_token = extract_token(reference)
    try:
        raw = run_lark_cli(
            [
                "lark-cli",
                "wiki",
                "spaces",
                "get_node",
                "--as",
                "user",
                "--params",
                json.dumps({"token": wiki_token}, ensure_ascii=False),
                "--format",
                "json",
            ]
        )
    except RuntimeError as exc:
        raise RuntimeError(
            f"failed to resolve wiki {context}; verify access and `{WIKI_READ_SCOPE}` scope, then retry"
        ) from exc

    try:
        return extract_wiki_node(raw)
    except RuntimeError as exc:
        raise RuntimeError(f"failed to resolve wiki {context}: {exc}") from exc


def resolve_source(args: argparse.Namespace, run_dir: Path) -> dict:
    if args.doc_url:
        resolved = {
            "input_kind": "doc_url",
            "resolved_kind": "wiki_url" if is_wiki_reference(args.doc_url) else "doc_url",
            "resolved_value": args.doc_url,
            "title": "",
            "search_candidates": [],
            "needs_user_choice": False,
        }
    elif args.doc_token:
        resolved = {
            "input_kind": "doc_token",
            "resolved_kind": "wiki_token" if is_wiki_token(args.doc_token) else "doc_token",
            "resolved_value": args.doc_token,
            "title": "",
            "search_candidates": [],
            "needs_user_choice": False,
        }
    elif args.doc_name:
        search_result = run_lark_cli(
            [
                "lark-cli",
                "docs",
                "+search",
                "--as",
                "user",
                "--format",
                "json",
                "--query",
                args.doc_name,
            ]
        )
        candidates = extract_search_candidates(search_result)
        if not candidates:
            raise RuntimeError(f"no document found for name: {args.doc_name}")
        if len(candidates) == 1:
            resolved = {
                "input_kind": "doc_name",
                **candidates[0],
                "search_candidates": candidates,
                "needs_user_choice": False,
            }
        else:
            resolved = {
                "input_kind": "doc_name",
                "resolved_kind": "",
                "resolved_value": "",
                "title": "",
                "search_candidates": candidates,
                "needs_user_choice": True,
            }
    else:
        raise RuntimeError("one of --doc-url, --doc-token, or --doc-name is required")

    write_json(run_dir / "resolved-source.json", resolved)
    return resolved


def choose_source_candidate(resolved_source_path: Path, candidate_index: int) -> dict:
    resolved = read_json(resolved_source_path)
    if not isinstance(resolved, dict):
        raise ValueError("resolved-source must be a JSON object")
    if not resolved.get("needs_user_choice"):
        raise RuntimeError("resolved source does not require an explicit user choice")
    candidates = resolved.get("search_candidates")
    if not isinstance(candidates, list) or not candidates:
        raise RuntimeError("resolved source is missing search_candidates")
    if candidate_index < 1 or candidate_index > len(candidates):
        raise ValueError(f"candidate-index must be between 1 and {len(candidates)}")

    candidate = candidates[candidate_index - 1]
    if not isinstance(candidate, dict):
        raise RuntimeError("selected search candidate must be a JSON object")
    resolved_kind = candidate.get("resolved_kind")
    resolved_value = candidate.get("resolved_value")
    if not resolved_kind or not resolved_value:
        raise RuntimeError("selected search candidate is missing resolved_kind or resolved_value")

    updated = dict(resolved)
    updated["resolved_kind"] = resolved_kind
    updated["resolved_value"] = resolved_value
    updated["title"] = candidate.get("title") or ""
    updated["needs_user_choice"] = False
    updated["selected_candidate_index"] = candidate_index
    write_json(resolved_source_path, updated)
    return updated


def resolve_fetch_target(resolved: dict) -> tuple[str, dict | None]:
    fetch_target = resolved["resolved_value"]
    wiki_node: dict | None = None

    if resolved.get("resolved_kind") in WIKI_RESOLVED_KINDS or is_wiki_reference(fetch_target) or is_wiki_token(fetch_target):
        wiki_node = resolve_wiki_node(fetch_target, context="source")
        obj_type = str(wiki_node.get("obj_type") or "").lower()
        if obj_type not in FETCHABLE_WIKI_OBJECT_TYPES:
            raise RuntimeError(
                "wiki source resolves to unsupported obj_type "
                f"'{wiki_node.get('obj_type')}'; supported source types: doc, docx"
            )
        obj_token = wiki_node.get("obj_token")
        if not obj_token:
            raise RuntimeError(f"wiki source resolved to '{obj_type or 'unknown'}' but did not include obj_token")
        fetch_target = obj_token

    return fetch_target, wiki_node


def ensure_resolved_source_ready(resolved: dict) -> None:
    if resolved.get("needs_user_choice"):
        raise RuntimeError("resolved source still requires explicit user choice before fetch")
    if not resolved.get("resolved_kind"):
        raise RuntimeError("resolved source is missing resolved_kind")
    if not resolved.get("resolved_value"):
        raise RuntimeError("resolved source is missing resolved_value")


def fetch_source(resolved: dict, run_dir: Path) -> dict:
    ensure_resolved_source_ready(resolved)
    offset = 0
    limit = 200
    pages: list[dict] = []
    markdown_parts: list[str] = []
    title = resolved.get("title") or ""
    fetch_target, wiki_node = resolve_fetch_target(resolved)
    seen_offsets: set[int] = set()

    while True:
        if offset in seen_offsets:
            raise RuntimeError(f"fetch pagination did not advance; repeated offset {offset}")
        seen_offsets.add(offset)
        raw_page = run_lark_cli(
            [
                "lark-cli",
                "docs",
                "+fetch",
                "--as",
                "user",
                "--format",
                "json",
                "--doc",
                fetch_target,
                "--offset",
                str(offset),
                "--limit",
                str(limit),
            ]
        )
        page = extract_fetch_payload(raw_page)
        pages.append(raw_page)
        if page.get("title") and not title:
            title = page["title"]
        markdown = page.get("markdown", "")
        if markdown:
            markdown_parts.append(markdown.rstrip())
        if not page.get("has_more"):
            break
        if "next_offset" not in page:
            raise RuntimeError("fetch pagination did not provide next_offset while has_more=true")
        next_offset = int(page["next_offset"])
        if next_offset <= offset:
            raise RuntimeError(
                f"fetch pagination did not advance; next_offset {next_offset} <= current offset {offset}"
            )
        offset = next_offset

    result = {
        "title": title or (wiki_node.get("title") if wiki_node else ""),
        "markdown": "\n\n".join(part for part in markdown_parts if part),
        "pages": len(pages),
        "raw_pages": pages,
        "resolved_fetch_target": fetch_target,
    }
    if wiki_node:
        result["wiki_node"] = wiki_node

    write_json(run_dir / "source.json", result)
    (run_dir / "source.md").write_text(
        result["markdown"] + ("\n" if result["markdown"] else ""),
        encoding="utf-8",
    )
    return result


def validate_outline(outline: dict) -> None:
    presentation = outline.get("presentation")
    if not isinstance(presentation, dict):
        raise ValueError("missing presentation")
    source = presentation.get("source")
    if not isinstance(source, dict):
        raise ValueError("missing presentation.source")
    slides = outline.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("slides must be a non-empty list")

    for field in ("title", "target_mode", "content_mode"):
        if not presentation.get(field):
            raise ValueError(f"missing presentation.{field}")
    if not isinstance(presentation["title"], str):
        raise ValueError("presentation.title must be a string")
    validate_text_budget("presentation.title", presentation["title"], PRESENTATION_TITLE_UNIT_LIMIT)
    if presentation["target_mode"] not in TARGET_MODES:
        raise ValueError(f"invalid presentation.target_mode: {presentation['target_mode']}")
    if presentation["content_mode"] not in CONTENT_MODES:
        raise ValueError(f"invalid presentation.content_mode: {presentation['content_mode']}")
    theme = presentation.get("theme")
    if theme is not None and theme not in THEME_PRESETS:
        raise ValueError(f"invalid presentation.theme: {theme}")
    cover_style = presentation.get("cover_style")
    if cover_style is not None and cover_style not in COVER_STYLES:
        raise ValueError(f"invalid presentation.cover_style: {cover_style}")
    if "subtitle" in presentation and presentation["subtitle"] is not None and not isinstance(
        presentation["subtitle"], str
    ):
        raise ValueError("presentation.subtitle must be a string")
    if isinstance(presentation.get("subtitle"), str):
        validate_text_budget(
            "presentation.subtitle",
            presentation["subtitle"],
            PRESENTATION_SUBTITLE_UNIT_LIMIT,
        )
    max_slides = presentation.get("max_slides")
    if max_slides is not None and (not isinstance(max_slides, int) or max_slides <= 0):
        raise ValueError("presentation.max_slides must be a positive integer")

    for field in ("input_kind", "resolved_kind", "resolved_value"):
        if not source.get(field):
            raise ValueError(f"missing presentation.source.{field}")

    if max_slides is not None and len(slides) > max_slides:
        raise ValueError(f"slides exceed presentation.max_slides: {len(slides)} > {max_slides}")

    for index, slide in enumerate(slides, start=1):
        for field in ("no", "role", "section_divider", "title", "layout", "key_points"):
            if field not in slide:
                raise ValueError(f"slide {index} missing {field}")
        if not isinstance(slide["no"], int) or slide["no"] <= 0:
            raise ValueError(f"slide {index} no must be a positive integer")
        if not isinstance(slide["title"], str):
            raise ValueError(f"slide {index} title must be a string")
        if slide["layout"] not in VALID_LAYOUTS:
            raise ValueError(f"slide {index} invalid layout: {slide['layout']}")
        if not isinstance(slide["key_points"], list):
            raise ValueError(f"slide {index} key_points must be a list")
        if len(slide["key_points"]) > 5:
            raise ValueError(f"slide {index} exceeds 5 key points")
        title_budget = 160 if slide.get("role") == "cover" else SLIDE_TITLE_UNIT_LIMITS[slide["layout"]]
        validate_text_budget(
            f"slide {index} title",
            slide["title"],
            title_budget,
        )
        total_key_point_units = 0
        for point_index, point in enumerate(slide["key_points"], start=1):
            if not isinstance(point, str):
                raise ValueError(f"slide {index} key_points[{point_index}] must be a string")
            validate_text_budget(
                f"slide {index} key_points[{point_index}]",
                point,
                KEY_POINT_UNIT_LIMITS[slide["layout"]],
            )
            total_key_point_units += text_units(point)
        if total_key_point_units > TOTAL_KEY_POINT_UNIT_LIMITS[slide["layout"]]:
            raise ValueError(
                f"slide {index} total key point budget exceeded: "
                f"{total_key_point_units} > {TOTAL_KEY_POINT_UNIT_LIMITS[slide['layout']]}"
            )
        if "objective" in slide and slide["objective"] is not None and not isinstance(slide["objective"], str):
            raise ValueError(f"slide {index} objective must be a string")
        if "notes" in slide and slide["notes"] is not None and not isinstance(slide["notes"], str):
            raise ValueError(f"slide {index} notes must be a string")
        if "source_sections" in slide:
            source_sections = slide["source_sections"]
            if source_sections is not None and (
                not isinstance(source_sections, list)
                or not all(isinstance(item, str) for item in source_sections)
            ):
                raise ValueError(f"slide {index} source_sections must be a list of strings")
        if not isinstance(slide["section_divider"], bool):
            raise ValueError(f"slide {index} section_divider must be a boolean")
        if slide["section_divider"] and (slide["role"] != "section" or slide["layout"] != "title-only"):
            raise ValueError(
                f"slide {index} section divider must use role='section' and layout='title-only'"
            )
        if slide["role"] == "section" and (not slide["section_divider"] or slide["layout"] != "title-only"):
            raise ValueError(f"slide {index} role='section' requires section_divider=true and title-only layout")
        if slide["layout"] == "title-only" and slide.get("role") != "cover":
            subtitle = effective_title_only_subtitle(slide)
            if subtitle:
                validate_text_budget(
                    f"slide {index} title-only subtitle",
                    subtitle,
                    TITLE_ONLY_SUBTITLE_UNIT_LIMIT,
                )

    if presentation["target_mode"] == "append":
        for slide in slides:
            if slide.get("role") == "cover":
                raise ValueError("append mode cannot include cover slides; use a section divider instead")


def text_shape(x: int, y: int, width: int, height: int, text_type: str, inner_xml: str) -> str:
    return (
        f'<shape type="text" topLeftX="{x}" topLeftY="{y}" width="{width}" height="{height}">'
        f'<content textType="{text_type}">{inner_xml}</content>'
        "</shape>"
    )


def rect_shape(
    x: int,
    y: int,
    width: int,
    height: int,
    fill_color: str,
    *,
    border_color: str = "rgb(203,213,225)",
    border_width: int = 2,
) -> str:
    return (
        f'<shape type="rect" topLeftX="{x}" topLeftY="{y}" width="{width}" height="{height}">'
        f'<fill><fillColor color="{fill_color}"/></fill>'
        f'<border color="{border_color}" width="{border_width}"/>'
        "</shape>"
    )


def circle_shape(
    x: int,
    y: int,
    diameter: int,
    fill_color: str,
    *,
    border_color: str | None = None,
    border_width: int = 0,
) -> str:
    """Create a circle shape. x and y are center coordinates."""
    radius = diameter // 2
    # For ellipse, topLeftX and topLeftY are top-left corner of bounding box
    top_left_x = x - radius
    top_left_y = y - radius
    if border_color is None:
        border_color = fill_color
    return (
        f'<shape type="ellipse" topLeftX="{top_left_x}" topLeftY="{top_left_y}" width="{diameter}" height="{diameter}">'
        f'<fill><fillColor color="{fill_color}"/></fill>'
        f'<border color="{border_color}" width="{border_width}"/>'
        "</shape>"
    )


def rounded_rect_shape(
    x: int,
    y: int,
    width: int,
    height: int,
    fill_color: str,
    *,
    border_color: str = "rgb(203,213,225)",
    border_width: int = 2,
    corner_radius: int = 12,
) -> str:
    """Create a rounded rectangle shape."""
    return (
        f'<shape type="roundRect" topLeftX="{x}" topLeftY="{y}" width="{width}" height="{height}" cornerRadius="{corner_radius}">'
        f'<fill><fillColor color="{fill_color}"/></fill>'
        f'<border color="{border_color}" width="{border_width}"/>'
        "</shape>"
    )


def centered_text_shape(
    x: int,
    y: int,
    width: int,
    height: int,
    text: str,
    *,
    color: str,
    font_size: int = 18,
) -> str:
    return text_shape(
        x,
        y,
        width,
        height,
        "body",
        f'<p textAlign="center"><span color="{color}" fontSize="{font_size}">{escape(text)}</span></p>',
    )


def styled_paragraph(text: str, *, color: str, font_size: int, align: str | None = None) -> str:
    align_attr = f' textAlign="{align}"' if align else ""
    return f'<p{align_attr}><span color="{color}" fontSize="{font_size}">{escape(text)}</span></p>'


def parse_rgb_color(value: str) -> tuple[int, int, int]:
    match = RGB_COLOR_RE.fullmatch(value.strip())
    if not match:
        raise ValueError(f"unsupported color format: {value}")
    return tuple(int(component) for component in match.groups())


def is_light_color(value: str) -> bool:
    red, green, blue = parse_rgb_color(value)
    luminance = (0.2126 * red + 0.7152 * green + 0.0722 * blue) / 255
    return luminance >= 0.62


def color_distance_rgb(left: tuple[int, int, int], right: tuple[int, int, int]) -> int:
    return sum((component_left - component_right) ** 2 for component_left, component_right in zip(left, right))


def extract_fill_color_from_style_node(node: ET.Element) -> str | None:
    style = node.find("{*}style")
    if style is None:
        return None
    fill = style.find("{*}fill")
    if fill is None:
        return None
    color_node = fill.find("{*}fillColor")
    if color_node is None:
        return None
    raw_color = color_node.get("color")
    if not isinstance(raw_color, str):
        return None
    normalized = raw_color.strip()
    if not RGB_COLOR_RE.fullmatch(normalized):
        return None
    return normalized


def extract_theme_background_color(root: ET.Element) -> str | None:
    theme = root.find("{*}theme")
    if theme is None:
        return None
    background = theme.find("{*}background")
    if background is None:
        return None
    color_node = background.find("{*}fillColor")
    if color_node is None:
        return None
    raw_color = color_node.get("color")
    if not isinstance(raw_color, str):
        return None
    normalized = raw_color.strip()
    if not RGB_COLOR_RE.fullmatch(normalized):
        return None
    return normalized


def extract_first_slide_bg_color(xml_content: str) -> str | None:
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        return None
    for slide in root.findall(".//{*}slide"):
        color = extract_fill_color_from_style_node(slide)
        if color:
            return color
    return extract_theme_background_color(root)


def infer_theme_from_background(background_color: str) -> str | None:
    try:
        target_rgb = parse_rgb_color(background_color)
    except ValueError:
        return None

    best_distance: int | None = None
    best_theme = None
    for theme_name, theme in THEME_PRESETS.items():
        try:
            theme_rgb = parse_rgb_color(theme["slide_bg"])
        except ValueError:
            continue
        distance = color_distance_rgb(theme_rgb, target_rgb)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_theme = theme_name
    return best_theme


def infer_theme_from_xml_presentation(xml_presentation_id: str) -> str | None:
    raw = run_lark_cli(
        [
            "lark-cli",
            "slides",
            "xml_presentations",
            "get",
            "--as",
            "user",
            "--params",
            json.dumps({"xml_presentation_id": xml_presentation_id}, ensure_ascii=False),
            "--format",
            "json",
        ]
    )
    payload = extract_xml_presentation_payload(raw)
    content = payload.get("content")
    if not isinstance(content, str):
        raise RuntimeError("xml_presentation.get did not return presentation content")
    background_color = extract_first_slide_bg_color(content)
    if background_color is None:
        return None
    return infer_theme_from_background(background_color)


def infer_append_target_theme(run_dir: Path) -> str | None:
    try:
        resolved_target = load_resolved_target(run_dir)
    except RuntimeError:
        return None
    try:
        return infer_theme_from_xml_presentation(str(resolved_target["xml_presentation_id"]))
    except (RuntimeError, ValueError):
        return None


def cover_text_colors(theme: dict, *, title_background: str, body_background: str) -> tuple[str, str]:
    title_color = theme["title_color"] if is_light_color(title_background) else theme["cover_title_color"]
    body_color = theme["body_color"] if is_light_color(body_background) else theme["cover_body_color"]
    return title_color, body_color


def get_cover_subtitle(presentation: dict | None, slide: dict) -> str:
    if presentation and presentation.get("subtitle"):
        return str(presentation["subtitle"])
    if slide.get("objective"):
        return str(slide["objective"])
    key_points = slide.get("key_points") or []
    if key_points:
        return str(key_points[0])
    return ""


def cover_title_metrics(title: str, *, modern: bool) -> tuple[int, int]:
    if len(title) > 90:
        return (24, 152) if modern else (22, 132)
    if len(title) > 55:
        return (28, 132) if modern else (24, 112)
    return (32, 132) if modern else (28, 96)


def cover_subtitle_metrics(subtitle: str, *, modern: bool) -> tuple[int, int, int]:
    if len(subtitle) > 120:
        return (18, 116, 80) if modern else (18, 84, 72)
    if len(subtitle) > 72:
        return (18, 96, 64) if modern else (19, 72, 60)
    return (20, 84, 48) if modern else (20, 54, 54)


def bullets_xml(points: list[str], *, color: str, font_size: int = 20) -> str:
    if not points:
        return "<p></p>"
    return "".join(
        f'<ul><li><p><span color="{color}" fontSize="{font_size}">{escape(point)}</span></p></li></ul>'
        for point in points
    )


def numbered_bullets_xml(points: list[str], *, color: str, font_size: int = 20, badge_fill: str, badge_text: str) -> str:
    """生成带序号的要点列表"""
    if not points:
        return "<p></p>"
    items = []
    for i, point in enumerate(points, start=1):
        # 序号使用主题色，文本使用正文色
        items.append(
            f'<p>'
            f'<span color="{badge_fill}" fontSize="{font_size}" bold="true">{i}. </span>'
            f'<span color="{color}" fontSize="{font_size}">{escape(point)}</span>'
            f'</p>'
        )
    return "".join(items)


def wrap_slide(data_xml: str, theme: dict, note_text: str = "") -> str:
    note_xml = ""
    if note_text:
        note_xml = f'<note><content textType="body"><p>{escape(note_text)}</p></content></note>'
    return (
        f'<slide xmlns="{SML_NS}">'
        f'<style><fill><fillColor color="{theme["slide_bg"]}"/></fill></style>'
        f"<data>{data_xml}</data>"
        f"{note_xml}"
        "</slide>"
    )


def parse_int_attr(node: ET.Element, name: str) -> int | None:
    raw = node.get(name)
    if raw is None:
        return None
    try:
        return int(float(raw))
    except ValueError:
        return None


def xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def validate_layout_semantics(index: int, layout: str, shapes: list[dict], lines: list[dict]) -> list[str]:
    violations: list[str] = []
    title_shapes = [shape for shape in shapes if shape["shape_type"] == "text" and shape["text_type"] == "title"]
    body_shapes = [shape for shape in shapes if shape["shape_type"] == "text" and shape["text_type"] == "body"]
    large_panels = [
        shape
        for shape in shapes
        if shape["shape_type"] == "rect" and shape["width"] >= 300 and shape["height"] >= 180 and shape["y"] >= 150
    ]

    if layout == "title-only":
        if not title_shapes:
            violations.append(f"slide {index} title-only layout is missing a title text shape")
    elif layout in {"title-body", "bullets"}:
        if not title_shapes:
            violations.append(f"slide {index} {layout} layout is missing a title text shape")
        if not body_shapes:
            violations.append(f"slide {index} {layout} layout is missing a body text shape")
    elif layout in {"two-column", "comparison"}:
        if not title_shapes:
            violations.append(f"slide {index} {layout} layout is missing a title text shape")
        if len(large_panels) < 2:
            violations.append(f"slide {index} {layout} layout is missing its two content panels")
    elif layout == "timeline":
        timeline_rail = [
            shape
            for shape in shapes
            if shape["shape_type"] == "rect" and shape["width"] <= 12 and shape["height"] >= 44
        ]
        timeline_items = [
            shape
            for shape in shapes
            if shape["shape_type"] == "rect" and shape["width"] >= 600 and shape["height"] >= 24
        ]
        if not title_shapes:
            violations.append("slide {index} timeline layout is missing a title text shape".format(index=index))
        if not timeline_rail:
            violations.append(f"slide {index} timeline layout is missing a rail marker")
        if not timeline_items:
            violations.append(f"slide {index} timeline layout is missing timeline item cards")
    elif layout == "metrics":
        metric_cards = [
            shape
            for shape in shapes
            if shape["shape_type"] == "rect" and 180 <= shape["width"] <= 240 and 60 <= shape["height"] <= 100
        ]
        if not title_shapes:
            violations.append(f"slide {index} metrics layout is missing a title text shape")
        if not metric_cards:
            violations.append(f"slide {index} metrics layout is missing metric cards")

    return violations


def layout_warning(index: int, layout: str, severity: str, kind: str, message: str) -> dict:
    return {
        "slide_index": index,
        "layout": layout,
        "severity": severity,
        "kind": kind,
        "message": message,
    }


def collect_layout_density_warnings(index: int, layout: str, shapes: list[dict], max_right: int, max_bottom: int) -> list[dict]:
    warnings: list[dict] = []
    if not max_right or not max_bottom:
        return warnings

    bbox_ratio = (max_right * max_bottom) / (SLIDE_WIDTH * SLIDE_HEIGHT)
    content_layouts = {"title-body", "bullets", "two-column", "comparison", "timeline", "metrics"}
    if layout in content_layouts and bbox_ratio < 0.20:
        warnings.append(
            layout_warning(
                index,
                layout,
                "low",
                "density_sparse",
                f"slide {index} {layout} layout may be visually sparse (content footprint ratio {bbox_ratio:.2f})",
            )
        )
    if layout in content_layouts and bbox_ratio > 0.88:
        warnings.append(
            layout_warning(
                index,
                layout,
                "medium",
                "density_dense",
                f"slide {index} {layout} layout may be visually dense (content footprint ratio {bbox_ratio:.2f})",
            )
        )

    body_shapes = [shape for shape in shapes if shape["shape_type"] == "text" and shape["text_type"] == "body"]
    if layout in {"title-body", "bullets"} and len(body_shapes) == 1 and body_shapes[0]["height"] >= 220:
        warnings.append(
            layout_warning(
                index,
                layout,
                "medium",
                "overflow_risk",
                f"slide {index} {layout} layout uses a tall single body region; review for overflow risk",
            )
        )

    return warnings


def validate_rendered_slides(rendered: list[str], layouts: list[str]) -> dict:
    per_slide: list[dict] = []
    violations: list[str] = []
    warnings: list[str] = []

    for index, (slide_xml, layout) in enumerate(zip(rendered, layouts, strict=True), start=1):
        try:
            root = ET.fromstring(slide_xml)
        except ET.ParseError as exc:
            raise ValueError(f"rendered slide {index} is invalid XML: {exc}") from exc

        shapes: list[dict] = []
        lines: list[dict] = []
        max_right = 0
        max_bottom = 0
        for node in root.findall(".//{*}shape") + root.findall(".//{*}img") + root.findall(".//{*}icon"):
            x = parse_int_attr(node, "topLeftX")
            y = parse_int_attr(node, "topLeftY")
            width = parse_int_attr(node, "width")
            height = parse_int_attr(node, "height")
            if None in {x, y, width, height}:
                continue
            shape_type = str(node.get("type") or "")
            text_type = ""
            if shape_type == "text":
                content = node.find("{*}content")
                text_type = str(content.get("textType") or "") if content is not None else ""
            shapes.append(
                {
                    "shape_type": shape_type or xml_local_name(node.tag),
                    "text_type": text_type,
                    "x": x,
                    "y": y,
                    "width": width,
                    "height": height,
                }
            )
            max_right = max(max_right, x + width)
            max_bottom = max(max_bottom, y + height)
            if x < 0 or y < 0 or x + width > SLIDE_WIDTH or y + height > SLIDE_HEIGHT:
                violations.append(
                    f"slide {index} has out-of-bounds element {node.tag}: "
                    f"({x},{y},{width},{height}) exceeds {SLIDE_WIDTH}x{SLIDE_HEIGHT}"
                )

        for node in root.findall(".//{*}line"):
            start_x = parse_int_attr(node, "startX")
            start_y = parse_int_attr(node, "startY")
            end_x = parse_int_attr(node, "endX")
            end_y = parse_int_attr(node, "endY")
            if None in {start_x, start_y, end_x, end_y}:
                continue
            lines.append(
                {
                    "start_x": start_x,
                    "start_y": start_y,
                    "end_x": end_x,
                    "end_y": end_y,
                }
            )
            max_right = max(max_right, start_x, end_x)
            max_bottom = max(max_bottom, start_y, end_y)
            if min(start_x, end_x) < 0 or min(start_y, end_y) < 0 or max(start_x, end_x) > SLIDE_WIDTH or max(
                start_y, end_y
            ) > SLIDE_HEIGHT:
                violations.append(
                    f"slide {index} has out-of-bounds line: "
                    f"({start_x},{start_y})->({end_x},{end_y}) exceeds {SLIDE_WIDTH}x{SLIDE_HEIGHT}"
                )

        violations.extend(validate_layout_semantics(index, layout, shapes, lines))
        slide_warnings = collect_layout_density_warnings(index, layout, shapes, max_right, max_bottom)
        warnings.extend(slide_warnings)
        per_slide.append(
            {
                "index": index,
                "layout": layout,
                "max_right": max_right,
                "max_bottom": max_bottom,
                "warnings": slide_warnings,
            }
        )

    if violations:
        raise ValueError("render layout guard failed: " + "; ".join(violations))

    return {
        "slides_checked": len(rendered),
        "slide_size": {"width": SLIDE_WIDTH, "height": SLIDE_HEIGHT},
        "warnings": warnings,
        "per_slide": per_slide,
    }


def split_points(points: list[str]) -> tuple[list[str], list[str]]:
    midpoint = max(1, (len(points) + 1) // 2)
    return points[:midpoint], points[midpoint:]


def render_cover_editorial_slide(slide: dict, presentation: dict, theme: dict) -> str:
    subtitle = get_cover_subtitle(presentation, slide)
    title_font_size, title_height = cover_title_metrics(slide["title"], modern=False)
    subtitle_font_size, _, subtitle_height = cover_subtitle_metrics(subtitle, modern=False)
    data_xml = [
        rect_shape(
            96,
            96,
            188,
            10,
            theme["cover_band_fill"],
            border_color=theme["cover_band_fill"],
            border_width=0,
        ),
        text_shape(
            96,
            118,
            736,
            title_height,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=title_font_size),
        ),
    ]
    if subtitle:
        data_xml.append(
            text_shape(
                96,
                118 + title_height + 22,
                620,
                subtitle_height,
                "body",
                styled_paragraph(subtitle, color=theme["body_color"], font_size=subtitle_font_size),
            )
        )
    return wrap_slide("".join(data_xml), theme, slide.get("notes", ""))


def render_cover_modern_slide(slide: dict, presentation: dict, theme: dict) -> str:
    subtitle = get_cover_subtitle(presentation, slide)
    hero_fill = theme["cover_modern_hero_fill"]
    card_fill = theme["cover_modern_card_fill"]
    title_color, body_color = cover_text_colors(
        theme,
        title_background=hero_fill,
        body_background=card_fill,
    )
    title_font_size, title_height = cover_title_metrics(slide["title"], modern=True)
    subtitle_font_size, card_height, subtitle_height = cover_subtitle_metrics(subtitle, modern=True)
    data_xml = [
        rect_shape(
            80,
            88,
            18,
            288,
            theme["cover_band_fill"],
            border_color=theme["cover_band_fill"],
            border_width=0,
        ),
        rect_shape(
            120,
            96,
            628,
            182,
            hero_fill,
            border_color=hero_fill,
            border_width=0,
        ),
        text_shape(
            144,
            122,
            580,
            title_height,
            "title",
            styled_paragraph(slide["title"], color=title_color, font_size=title_font_size),
        ),
    ]
    if subtitle:
        data_xml.append(
            rect_shape(
                144,
                290,
                548,
                card_height,
                card_fill,
                border_color=theme["border_color"],
            )
        )
        data_xml.append(
            text_shape(
                168,
                308,
                500,
                subtitle_height,
                "body",
                styled_paragraph(subtitle, color=body_color, font_size=subtitle_font_size),
            )
        )
    return wrap_slide("".join(data_xml), theme, slide.get("notes", ""))


def render_plain_title_only_slide(slide: dict, theme: dict) -> str:
    subtitle = slide.get("objective") or (slide.get("key_points") or [""])[0]
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])

    data_xml = [
        # 左侧装饰条
        rect_shape(
            0,
            0,
            12,
            SLIDE_HEIGHT,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        ),
        # 顶部装饰线
        rect_shape(
            12,
            0,
            SLIDE_WIDTH - 12,
            4,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        ),
        # 标题
        text_shape(
            96,
            200,
            736,
            96,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=32),
        ),
    ]
    if subtitle:
        data_xml.append(
            text_shape(
                96,
                300,
                768,
                120,
                "body",
                styled_paragraph(subtitle, color=theme["body_color"], font_size=20),
            )
        )
    # 右下角页码
    data_xml.append(
        text_shape(
            SLIDE_WIDTH - 80,
            SLIDE_HEIGHT - 40,
            60,
            24,
            "body",
            styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
        )
    )
    return wrap_slide("".join(data_xml), theme, slide.get("notes", ""))


def render_section_divider_slide(slide: dict, theme: dict) -> str:
    subtitle = slide.get("objective") or (slide.get("key_points") or [""])[0]
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])
    panel_fill = theme.get("panel_fill", "rgb(248,250,252)")

    # 章节页使用更大更突出的设计
    data_xml = [
        # 左侧宽装饰条
        rect_shape(
            0,
            0,
            24,
            SLIDE_HEIGHT,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        ),
        # 右侧大装饰圆环
        circle_shape(SLIDE_WIDTH - 100, 400, 200, panel_fill),
        circle_shape(SLIDE_WIDTH - 100, 400, 160, theme["slide_bg"]),
        # 标题左侧装饰线
        rect_shape(
            80,
            200,
            120,
            8,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        ),
        # 装饰圆点
        circle_shape(60, 204, 12, badge_fill),
        # 标题
        text_shape(
            80,
            220,
            700,
            80,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=36),
        ),
    ]
    if subtitle:
        data_xml.append(
            text_shape(
                80,
                310,
                700,
                60,
                "body",
                styled_paragraph(subtitle, color=theme["muted_color"], font_size=20),
            )
        )
    # 右下角页码
    data_xml.append(
        text_shape(
            SLIDE_WIDTH - 80,
            SLIDE_HEIGHT - 40,
            60,
            24,
            "body",
            styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
        )
    )
    return wrap_slide("".join(data_xml), theme, slide.get("notes", ""))


def render_title_only_slide(
    slide: dict,
    theme: dict,
    cover_style: str = "editorial",
    presentation: dict | None = None,
) -> str:
    is_section_divider = bool(slide.get("section_divider")) or slide.get("role") == "section"
    if is_section_divider:
        return render_section_divider_slide(slide, theme)
    if slide.get("role") == "cover":
        if cover_style == "modern":
            return render_cover_modern_slide(slide, presentation or {}, theme)
        return render_cover_editorial_slide(slide, presentation or {}, theme)
    return render_plain_title_only_slide(slide, theme)


def render_title_body_slide(slide: dict, theme: dict) -> str:
    key_points = slide.get("key_points", [])
    # 判断是否使用卡片式布局（3个及以下要点）
    use_card_layout = len(key_points) <= 3 and len(key_points) > 0

    # 获取主题色
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])
    badge_text = theme.get("number_badge_text", "#ffffff")
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    title_bg = theme.get("title_bg_fill", theme["panel_fill"])

    if use_card_layout:
        # 卡片式布局：每个要点一个卡片，增加视觉层次
        card_width = 260
        card_height = 200
        card_gap = 24
        total_width = len(key_points) * card_width + (len(key_points) - 1) * card_gap
        start_x = (SLIDE_WIDTH - total_width) // 2

        data_xml = (
            # 左侧装饰条 - 更宽更有存在感
            rect_shape(
                0,
                0,
                12,
                SLIDE_HEIGHT,
                sidebar_accent,
                border_color=sidebar_accent,
                border_width=0,
            )
            # 顶部装饰线
            + rect_shape(
                12,
                0,
                SLIDE_WIDTH - 12,
                4,
                sidebar_accent,
                border_color=sidebar_accent,
                border_width=0,
            )
            # 标题背景 - 渐变效果区域
            + rect_shape(24, 60, SLIDE_WIDTH - 48, 80, title_bg, border_color=theme["border_color"])
            # 标题左侧装饰点
            + circle_shape(40, 100, 12, badge_fill)
            # 标题
            + text_shape(
                64,
                84,
                SLIDE_WIDTH - 100,
                48,
                "title",
                styled_paragraph(slide["title"], color=theme["title_color"], font_size=28),
            )
            # 右下角页码装饰
            + text_shape(
                SLIDE_WIDTH - 80,
                SLIDE_HEIGHT - 40,
                60,
                24,
                "body",
                styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
            )
        )
        # 卡片 - 带阴影效果和序号圆圈
        for i, point in enumerate(key_points, start=1):
            card_x = start_x + (i - 1) * (card_width + card_gap)
            # 卡片阴影（通过偏移的灰色矩形模拟）
            data_xml += rect_shape(card_x + 3, 183, card_width, card_height, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
            # 卡片主体
            data_xml += (
                rect_shape(card_x, 180, card_width, card_height, theme["panel_fill"], border_color=theme["border_color"])
                # 序号圆圈
                + circle_shape(card_x + card_width // 2, 200, 32, badge_fill)
                + text_shape(
                    card_x + card_width // 2 - 8,
                    191,
                    16,
                    24,
                    "body",
                    styled_paragraph(str(i), color=badge_text, font_size=18),
                )
                # 要点文本
                + text_shape(
                    card_x + 20,
                    250,
                    card_width - 40,
                    card_height - 80,
                    "body",
                    f'<p align="center"><span color="{theme["body_color"]}" fontSize="16">{escape(point)}</span></p>',
                )
            )
    else:
        # 列表布局：带左侧装饰条和序号
        data_xml = (
            # 左侧装饰条 - 更宽
            rect_shape(
                0,
                0,
                12,
                SLIDE_HEIGHT,
                sidebar_accent,
                border_color=sidebar_accent,
                border_width=0,
            )
            # 顶部装饰线
            + rect_shape(
                12,
                0,
                SLIDE_WIDTH - 12,
                4,
                sidebar_accent,
                border_color=sidebar_accent,
                border_width=0,
            )
            # 标题背景
            + rect_shape(24, 60, SLIDE_WIDTH - 48, 80, title_bg, border_color=theme["border_color"])
            # 标题左侧装饰点
            + circle_shape(40, 100, 12, badge_fill)
            # 标题
            + text_shape(
                64,
                84,
                SLIDE_WIDTH - 100,
                48,
                "title",
                styled_paragraph(slide["title"], color=theme["title_color"], font_size=28),
            )
            # 内容卡片
            + rect_shape(24, 160, SLIDE_WIDTH - 48, 320, theme["panel_fill"], border_color=theme["border_color"])
            # 序号列表
            + text_shape(
                48,
                180,
                SLIDE_WIDTH - 96,
                280,
                "body",
                numbered_bullets_xml(key_points, color=theme["body_color"], font_size=20, badge_fill=badge_fill, badge_text=badge_text),
            )
            # 右下角页码
            + text_shape(
                SLIDE_WIDTH - 80,
                SLIDE_HEIGHT - 40,
                60,
                24,
                "body",
                styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
            )
        )
    return wrap_slide(data_xml, theme, slide.get("notes", ""))


def render_two_column_slide(slide: dict, theme: dict) -> str:
    left_points, right_points = split_points(slide.get("key_points", []))
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    title_bg = theme.get("title_bg_fill", theme["panel_fill"])
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])
    badge_text = theme.get("number_badge_text", "#ffffff")

    data_xml = (
        # 左侧装饰条 - 全高
        rect_shape(
            0,
            0,
            12,
            SLIDE_HEIGHT,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 顶部装饰线
        + rect_shape(
            12,
            0,
            SLIDE_WIDTH - 12,
            4,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 标题区域背景
        + rect_shape(24, 60, SLIDE_WIDTH - 48, 80, title_bg, border_color=theme["border_color"])
        # 标题左侧装饰圆点
        + circle_shape(44, 100, 12, badge_fill)
        # 标题
        + text_shape(
            68,
            84,
            SLIDE_WIDTH - 100,
            48,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=28),
        )
        # 左栏卡片 - 带阴影
        + rect_shape(63, 183, 400, 280, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
        + rect_shape(60, 180, 400, 280, theme["panel_fill"], border_color=theme["border_color"])
        # 右栏卡片 - 带阴影
        + rect_shape(503, 183, 400, 280, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
        + rect_shape(500, 180, 400, 280, theme["panel_alt_fill"], border_color=theme["border_color"])
        # 左栏内容
        + text_shape(84, 200, 352, 244, "body", numbered_bullets_xml(left_points, color=theme["body_color"], font_size=18, badge_fill=badge_fill, badge_text=badge_text))
        # 右栏内容
        + text_shape(524, 200, 352, 244, "body", numbered_bullets_xml(right_points, color=theme["body_color"], font_size=18, badge_fill=badge_fill, badge_text=badge_text))
        # 右下角页码
        + text_shape(
            SLIDE_WIDTH - 80,
            SLIDE_HEIGHT - 40,
            60,
            24,
            "body",
            styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
        )
    )
    return wrap_slide(data_xml, theme, slide.get("notes", ""))


def render_comparison_slide(slide: dict, theme: dict) -> str:
    # 尝试从标题解析对比标签（如 "A vs B"、"A/B"、"A 与 B"）
    title = slide.get("title", "")
    left_label = "方案 A"
    right_label = "方案 B"

    # 尝试解析标题中的对比关系
    for sep in [" vs ", " VS ", " Vs ", " vs. ", " / ", " 与 ", " 和 "]:
        if sep in title:
            parts = title.split(sep, 1)
            if len(parts) == 2:
                left_label = parts[0].strip()
                right_label = parts[1].strip()
                break

    # 也可以从 source_sections 覆盖
    labels = slide.get("source_sections") or []
    if len(labels) >= 2:
        left_label = labels[0]
        right_label = labels[1]

    left_points, right_points = split_points(slide.get("key_points", []))
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    title_bg = theme.get("title_bg_fill", theme["panel_fill"])
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])
    badge_text_color = theme.get("number_badge_text", "#ffffff")

    data_xml = (
        # 左侧装饰条 - 全高
        rect_shape(
            0,
            0,
            12,
            SLIDE_HEIGHT,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 顶部装饰线
        + rect_shape(
            12,
            0,
            SLIDE_WIDTH - 12,
            4,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 标题区域背景
        + rect_shape(24, 60, SLIDE_WIDTH - 48, 80, title_bg, border_color=theme["border_color"])
        # 标题左侧装饰圆点
        + circle_shape(44, 100, 12, badge_fill)
        # 标题
        + text_shape(
            68,
            84,
            SLIDE_WIDTH - 100,
            48,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=28),
        )
        # 中间分隔装饰线
        + rect_shape(
            478,
            170,
            4,
            280,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 左栏卡片 - 带阴影
        + rect_shape(63, 173, 400, 280, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
        + rect_shape(60, 170, 400, 280, theme["comparison_left_fill"], border_color=theme["border_color"])
        # 右栏卡片 - 带阴影
        + rect_shape(499, 173, 400, 280, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
        + rect_shape(496, 170, 400, 280, theme["comparison_right_fill"], border_color=theme["border_color"])
        # 左栏标签背景
        + rect_shape(60, 170, 400, 36, badge_fill, border_color=badge_fill, border_width=0)
        # 左栏标签文本
        + text_shape(
            80,
            176,
            360,
            28,
            "body",
            styled_paragraph(left_label, color=badge_text_color, font_size=16),
        )
        # 右栏标签背景
        + rect_shape(496, 170, 400, 36, badge_fill, border_color=badge_fill, border_width=0)
        # 右栏标签文本
        + text_shape(
            516,
            176,
            360,
            28,
            "body",
            styled_paragraph(right_label, color=badge_text_color, font_size=16),
        )
        # 左栏内容
        + text_shape(80, 220, 360, 220, "body", bullets_xml(left_points, color=theme["body_color"], font_size=18))
        # 右栏内容
        + text_shape(516, 220, 360, 220, "body", bullets_xml(right_points, color=theme["body_color"], font_size=18))
        # 右下角页码
        + text_shape(
            SLIDE_WIDTH - 80,
            SLIDE_HEIGHT - 40,
            60,
            24,
            "body",
            styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
        )
    )
    return wrap_slide(data_xml, theme, slide.get("notes", ""))


def render_timeline_slide(slide: dict, theme: dict) -> str:
    key_points = slide.get("key_points", [])
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    title_bg = theme.get("title_bg_fill", theme["panel_fill"])
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])

    base_y = 170
    rail_height = max(44, len(key_points) * 62)

    data_xml = (
        # 左侧装饰条
        rect_shape(
            0,
            0,
            12,
            SLIDE_HEIGHT,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 顶部装饰线
        + rect_shape(
            12,
            0,
            SLIDE_WIDTH - 12,
            4,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 标题背景
        + rect_shape(24, 60, SLIDE_WIDTH - 48, 80, title_bg, border_color=theme["border_color"])
        # 标题装饰圆点
        + circle_shape(44, 100, 12, badge_fill)
        # 标题
        + text_shape(
            68,
            84,
            SLIDE_WIDTH - 100,
            48,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=28),
        )
        # 时间轴垂直线
        + rect_shape(
            100,
            base_y,
            6,
            rail_height,
            theme["timeline_rail_fill"],
            border_color=theme["timeline_rail_fill"],
            border_width=0,
        )
    )

    # 时间轴项目 - 带圆点标记和卡片
    for index, point in enumerate(key_points):
        y = base_y + index * 62
        # 时间轴圆点
        data_xml += circle_shape(103, y + 22, 16, badge_fill)
        # 时间轴卡片阴影
        data_xml += rect_shape(127, y + 3, 720, 44, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
        # 时间轴卡片
        data_xml += rect_shape(124, y, 720, 44, theme["timeline_item_fill"], border_color=theme["border_color"])
        # 内容文本
        data_xml += text_shape(
            140,
            y + 10,
            688,
            28,
            "body",
            styled_paragraph(point, color=theme["body_color"], font_size=16),
        )

    # 右下角页码
    data_xml += text_shape(
        SLIDE_WIDTH - 80,
        SLIDE_HEIGHT - 40,
        60,
        24,
        "body",
        styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
    )
    return wrap_slide(data_xml, theme, slide.get("notes", ""))


def render_metrics_slide(slide: dict, theme: dict) -> str:
    key_points = slide.get("key_points", [])
    sidebar_accent = theme.get("sidebar_accent", theme["content_band_fill"])
    title_bg = theme.get("title_bg_fill", theme["panel_fill"])
    badge_fill = theme.get("number_badge_fill", theme["content_band_fill"])

    data_xml = (
        # 左侧装饰条
        rect_shape(
            0,
            0,
            12,
            SLIDE_HEIGHT,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 顶部装饰线
        + rect_shape(
            12,
            0,
            SLIDE_WIDTH - 12,
            4,
            sidebar_accent,
            border_color=sidebar_accent,
            border_width=0,
        )
        # 标题背景
        + rect_shape(24, 60, SLIDE_WIDTH - 48, 80, title_bg, border_color=theme["border_color"])
        # 标题装饰圆点
        + circle_shape(44, 100, 12, badge_fill)
        # 标题
        + text_shape(
            68,
            84,
            SLIDE_WIDTH - 100,
            48,
            "title",
            styled_paragraph(slide["title"], color=theme["title_color"], font_size=28),
        )
    )

    # 指标卡片 - 带阴影和数字标记
    for index, point in enumerate(key_points):
        row = index // 3
        col = index % 3
        x = 60 + col * 280
        y = 180 + row * 120
        fill_color = theme["metrics_highlight_fill"] if index == 0 else theme["metrics_card_fill"]
        # 卡片阴影
        data_xml += rect_shape(x + 3, y + 3, 240, 88, "rgb(200,200,200)", border_color="rgb(200,200,200)", border_width=0)
        # 卡片主体
        data_xml += rect_shape(x, y, 240, 88, fill_color, border_color=theme["border_color"])
        # 序号圆圈
        data_xml += circle_shape(x + 24, y + 44, 24, badge_fill)
        data_xml += text_shape(
            x + 16,
            y + 36,
            16,
            20,
            "body",
            styled_paragraph(str(index + 1), color=theme.get("number_badge_text", "#ffffff"), font_size=14),
        )
        # 指标文本
        data_xml += text_shape(
            x + 44,
            y + 28,
            180,
            40,
            "body",
            styled_paragraph(point, color=theme["body_color"], font_size=16),
        )

    # 右下角页码
    data_xml += text_shape(
        SLIDE_WIDTH - 80,
        SLIDE_HEIGHT - 40,
        60,
        24,
        "body",
        styled_paragraph(str(slide.get("no", "")), color=theme["muted_color"], font_size=14),
    )
    return wrap_slide(data_xml, theme, slide.get("notes", ""))


RENDERERS = {
    "title-only": render_title_only_slide,
    "title-body": render_title_body_slide,
    "two-column": render_two_column_slide,
    "bullets": render_title_body_slide,
    "comparison": render_comparison_slide,
    "timeline": render_timeline_slide,
    "metrics": render_metrics_slide,
}


def render_outline(outline: dict, run_dir: Path) -> dict:
    validate_outline(outline)
    presentation = outline["presentation"]
    presentation_theme = presentation.get("theme")
    inferred_theme = None
    if presentation_theme is None and presentation["target_mode"] == "append":
        inferred_theme = infer_append_target_theme(run_dir)
        if inferred_theme is not None:
            presentation_theme = inferred_theme
    theme_name = presentation_theme or DEFAULT_THEME_BY_CONTENT_MODE[presentation["content_mode"]]
    cover_style = outline["presentation"].get("cover_style") or DEFAULT_COVER_STYLE_BY_THEME[theme_name]
    theme = THEME_PRESETS[theme_name]
    rendered = []
    rendered_layouts: list[str] = []
    for slide in outline["slides"]:
        rendered_layouts.append(slide["layout"])
        renderer = RENDERERS.get(slide["layout"])
        if renderer is None:
            raise ValueError(f"unsupported render layout: {slide['layout']}")
        if slide["layout"] == "title-only":
            rendered.append(
                render_title_only_slide(
                    slide,
                    theme,
                    cover_style=cover_style,
                    presentation=outline["presentation"],
                )
            )
        elif slide["layout"] in {"title-body", "bullets"}:
            rendered.append(render_title_body_slide(slide, theme))
        else:
            rendered.append(renderer(slide, theme))

    outline_fingerprint = fingerprint_payload(outline)
    slides_fingerprint = fingerprint_payload(rendered)
    layout_guard = validate_rendered_slides(rendered, rendered_layouts)
    result = {"slides": rendered, "count": len(rendered)}
    write_json(
        run_dir / "render-summary.json",
        {
            "count": len(rendered),
            "layouts": rendered_layouts,
            "theme": theme_name,
            "append_inferred_theme": inferred_theme,
            "cover_style": cover_style,
            "layout_guard": layout_guard,
            "outline_fingerprint": outline_fingerprint,
            "slides_fingerprint": slides_fingerprint,
        },
    )
    write_json(run_dir / "slides.json", rendered)
    return result


def extract_create_payload(raw: dict) -> dict:
    data = raw.get("data")
    if isinstance(data, dict) and data.get("xml_presentation_id"):
        return data
    if raw.get("xml_presentation_id"):
        return raw
    raise RuntimeError("unexpected slides +create response shape")


def extract_slide_create_payload(raw: dict) -> dict:
    if raw.get("slide_id"):
        return raw
    data = raw.get("data")
    if isinstance(data, dict) and data.get("slide_id"):
        return data
    raise RuntimeError("unexpected xml_presentation.slide create response shape")


def normalize_publish_result(
    target_mode: str,
    xml_presentation_id: str,
    url: str | None,
    slide_ids: list[str],
    run_dir: Path,
) -> dict:
    return {
        "target_mode": target_mode,
        "xml_presentation_id": xml_presentation_id,
        "url": url,
        "slide_ids": slide_ids,
        "slides_added": len(slide_ids),
        "run_dir": str(run_dir),
    }


def ensure_render_consistency(outline: dict, slides: list[str], run_dir: Path) -> None:
    summary_path = run_dir / "render-summary.json"
    if not summary_path.exists():
        raise RuntimeError("render-summary.json is required before publish")
    summary = read_json(summary_path)
    if not isinstance(summary, dict):
        raise RuntimeError("render-summary.json must be a JSON object")
    if summary.get("count") != len(slides):
        raise RuntimeError("slides.json does not match render-summary count")
    if summary.get("outline_fingerprint") != fingerprint_payload(outline):
        raise RuntimeError("outline.json no longer matches the rendered slides; rerun render")
    if summary.get("slides_fingerprint") != fingerprint_payload(slides):
        raise RuntimeError("slides.json no longer matches render-summary; rerun render")


def normalize_optional_url(value: str | None) -> str | None:
    if value is None:
        return None
    normalized = value.strip()
    return normalized or None


def resolve_slides_url(slides_url: str, *, context: str) -> tuple[str, str | None]:
    if "/slides/" in slides_url:
        return extract_token(slides_url), slides_url
    if "/wiki/" in slides_url:
        node = resolve_wiki_node(slides_url, context=context)
        obj_type = str(node.get("obj_type") or "").lower()
        if obj_type != "slides":
            title = node.get("title")
            title_prefix = f" '{title}'" if title else ""
            raise RuntimeError(
                f"{context} wiki node{title_prefix} resolves to obj_type '{node.get('obj_type')}', not slides"
            )
        obj_token = node.get("obj_token")
        if not obj_token:
            raise RuntimeError(f"{context} wiki node resolved to slides but did not include obj_token")
        return obj_token, slides_url
    raise RuntimeError(f"unsupported {context}_slides_url; expected a /slides/ or /wiki/ URL")


def resolve_target_slides_url(target_slides_url: str) -> tuple[str, str | None]:
    return resolve_slides_url(target_slides_url, context="target")


def resolve_target(target_slides_url: str | None, run_dir: Path) -> dict:
    normalized_target = normalize_optional_url(target_slides_url)
    if normalized_target is None:
        raise ValueError("resolve-target requires target_slides_url")
    presentation_id, url = resolve_target_slides_url(normalized_target)
    result = {
        "target_mode": "append",
        "preflight_kind": "target_append",
        "requested_url": normalized_target,
        "xml_presentation_id": presentation_id,
        "url": url,
    }
    write_json(run_dir / "resolved-target.json", result)
    return result


def load_resolved_target(run_dir: Path) -> dict:
    path = run_dir / "resolved-target.json"
    if not path.exists():
        raise RuntimeError("resolved-target.json is required before append publish")
    resolved = read_json(path)
    if not isinstance(resolved, dict):
        raise RuntimeError("resolved-target.json must be a JSON object")
    if resolved.get("target_mode") != "append":
        raise RuntimeError("resolved-target.json must describe an append target")
    xml_presentation_id = resolved.get("xml_presentation_id")
    if not isinstance(xml_presentation_id, str) or not xml_presentation_id:
        raise RuntimeError("resolved-target.json missing xml_presentation_id")
    url = resolved.get("url")
    if url is not None and not isinstance(url, str):
        raise RuntimeError("resolved-target.json url must be a string when present")
    requested_url = resolved.get("requested_url")
    if requested_url is not None and not isinstance(requested_url, str):
        raise RuntimeError("resolved-target.json requested_url must be a string when present")
    return resolved


def resolve_publish_append_target(run_dir: Path, target_slides_url: str | None) -> dict:
    resolved = load_resolved_target(run_dir)
    reference_url = normalize_optional_url(target_slides_url)
    if reference_url is None:
        return resolved

    resolved_presentation_id = str(resolved["xml_presentation_id"])
    reference_presentation_id, _ = resolve_target_slides_url(reference_url)
    if reference_presentation_id != resolved_presentation_id:
        raise RuntimeError("target_slides_url does not match resolved-target.json")
    return resolved


def create_slide_in_presentation(presentation_id: str, slide_xml: str) -> str:
    raw = run_lark_cli(
        [
            "lark-cli",
            "slides",
            "xml_presentation.slide",
            "create",
            "--as",
            "user",
            "--params",
            json.dumps({"xml_presentation_id": presentation_id}, ensure_ascii=False),
            "--data",
            json.dumps({"slide": {"content": slide_xml}}, ensure_ascii=False),
            "--format",
            "json",
        ]
    )
    return str(extract_slide_create_payload(raw)["slide_id"])


def publish_new_deck(title: str, slides: list[str], run_dir: Path) -> dict:
    create_raw = run_lark_cli(
        [
            "lark-cli",
            "slides",
            "+create",
            "--as",
            "user",
            "--title",
            title,
        ]
    )
    payload = extract_create_payload(create_raw)
    presentation_id = str(payload["xml_presentation_id"])
    url = payload.get("url")
    slide_ids: list[str] = []
    try:
        for slide_xml in slides:
            slide_ids.append(create_slide_in_presentation(presentation_id, slide_xml))
    except Exception as exc:
        raise PublishError(
            f"publish failed after {len(slide_ids)} slides: {exc}",
            normalize_publish_result("new", presentation_id, url, slide_ids, run_dir),
        ) from exc

    return normalize_publish_result("new", presentation_id, url, slide_ids, run_dir)


def publish_append(resolved_target: dict, slides: list[str], run_dir: Path) -> dict:
    presentation_id = str(resolved_target["xml_presentation_id"])
    url = resolved_target.get("url")
    slide_ids: list[str] = []
    try:
        for slide_xml in slides:
            slide_ids.append(create_slide_in_presentation(presentation_id, slide_xml))
    except Exception as exc:
        raise PublishError(
            f"append failed after {len(slide_ids)} slides: {exc}",
            normalize_publish_result("append", presentation_id, url, slide_ids, run_dir),
        ) from exc
    return normalize_publish_result("append", presentation_id, url, slide_ids, run_dir)


def publish_slides(
    outline: dict,
    slides: list[str],
    run_dir: Path,
    target_slides_url: str | None,
) -> dict:
    validate_outline(outline)
    ensure_render_consistency(outline, slides, run_dir)
    target_mode = outline["presentation"]["target_mode"]
    has_target_reference = normalize_optional_url(target_slides_url) is not None
    try:
        if target_mode == "new":
            if has_target_reference:
                raise ValueError("publish with target_mode=new does not accept target_slides_url")
            result = publish_new_deck(outline["presentation"]["title"], slides, run_dir)
        elif target_mode == "append":
            result = publish_append(
                resolve_publish_append_target(run_dir, target_slides_url),
                slides,
                run_dir,
            )
        else:
            raise RuntimeError(f"unsupported target_mode: {target_mode}")
    except PublishError as exc:
        write_json(run_dir / "publish-result.json", exc.result)
        raise

    write_json(run_dir / "publish-result.json", result)
    return result


def load_slides_json(path: Path) -> list[str]:
    payload = read_json(path)
    if not isinstance(payload, list):
        raise ValueError("slides-json must contain a JSON array of slide XML strings")
    if not all(isinstance(item, str) for item in payload):
        raise ValueError("slides-json entries must be strings")
    return payload


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    if args.command == "resolve-source":
        sources = [value for value in (args.doc_url, args.doc_token, args.doc_name) if value]
        if len(sources) != 1:
            raise SystemExit(2)
        resolve_source(args, ensure_run_dir(args.run_dir))
        return 0

    if args.command == "resolve-target":
        resolve_target(args.target_slides_url, ensure_run_dir(args.run_dir))
        return 0

    if args.command == "choose-source":
        choose_source_candidate(Path(args.resolved_source), args.candidate_index)
        return 0

    if args.command == "fetch":
        run_dir = ensure_run_dir(args.run_dir)
        resolved = read_json(Path(args.resolved_source))
        if not isinstance(resolved, dict):
            raise ValueError("resolved-source must be a JSON object")
        fetch_source(resolved, run_dir)
        return 0

    if args.command == "validate-outline":
        try:
            outline = read_json(Path(args.outline))
            if not isinstance(outline, dict):
                raise ValueError("outline must be a JSON object")
            validate_outline(outline)
            return 0
        except (ValueError, RuntimeError, json.JSONDecodeError) as exc:
            return emit_error("validate-outline", exc)

    if args.command == "render":
        run_dir = ensure_run_dir(args.run_dir)
        outline = read_json(Path(args.outline))
        if not isinstance(outline, dict):
            raise ValueError("outline must be a JSON object")
        render_outline(outline, run_dir)
        return 0

    if args.command == "publish":
        run_dir = ensure_run_dir(args.run_dir)
        outline = read_json(Path(args.outline))
        if not isinstance(outline, dict):
            raise ValueError("outline must be a JSON object")
        slides = load_slides_json(Path(args.slides_json))
        publish_slides(
            outline,
            slides,
            run_dir,
            args.target_slides_url,
        )
        return 0

    raise RuntimeError(f"unsupported command: {args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
