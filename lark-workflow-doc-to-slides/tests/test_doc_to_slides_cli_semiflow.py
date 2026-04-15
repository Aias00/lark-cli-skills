from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "doc_to_slides.py"


class CliSemiflowTests(unittest.TestCase):
    def write_fake_lark_cli(self, root: Path) -> tuple[Path, Path, Path]:
        bin_dir = root / "bin"
        bin_dir.mkdir(parents=True, exist_ok=True)
        log_path = root / "lark-cli.log"
        state_path = root / "lark-cli-state.json"
        script_path = bin_dir / "lark-cli"
        script_path.write_text(
            textwrap.dedent(
                """\
                #!/usr/bin/env python3
                import json
                import os
                import sys
                from pathlib import Path

                argv = sys.argv[1:]
                log_path = Path(os.environ["FAKE_LARK_CLI_LOG"])
                state_path = Path(os.environ["FAKE_LARK_CLI_STATE"])
                state = {"slide_count": 0}
                if state_path.exists():
                    state = json.loads(state_path.read_text(encoding="utf-8"))
                with log_path.open("a", encoding="utf-8") as fh:
                    fh.write(json.dumps(argv, ensure_ascii=False) + "\\n")

                def arg_value(flag: str) -> str:
                    index = argv.index(flag)
                    return argv[index + 1]

                if argv[:2] == ["docs", "+search"]:
                    print(json.dumps({
                        "data": {
                            "results": [
                                {
                                    "title": "Quarterly Review",
                                    "entity_type": "DOCX",
                                    "result_meta": {
                                        "title": "Quarterly Review",
                                        "url": "https://example.feishu.cn/docx/doccn123"
                                    }
                                },
                                {
                                    "title": "Quarterly Review Wiki",
                                    "entity_type": "WIKI",
                                    "result_meta": {
                                        "title": "Quarterly Review Wiki",
                                        "url": "https://example.feishu.cn/wiki/wiki_token"
                                    }
                                }
                            ]
                        }
                    }, ensure_ascii=False))
                    sys.exit(0)

                if argv[:3] == ["wiki", "spaces", "get_node"]:
                    token = json.loads(arg_value("--params"))["token"]
                    payload = {
                        "data": {
                            "node": {
                                "obj_type": "docx",
                                "obj_token": "doxcnResolved123",
                                "title": "Quarterly Review Wiki"
                            }
                        }
                    }
                    if token == "target_wiki":
                        payload = {
                            "data": {
                                "node": {
                                    "obj_type": "slides",
                                    "obj_token": "sldcnTarget456",
                                    "title": "Quarterly Review Deck"
                                }
                            }
                        }
                    print(json.dumps(payload, ensure_ascii=False))
                    sys.exit(0)

                if argv[:2] == ["docs", "+fetch"]:
                    offset = arg_value("--offset")
                    if offset == "0":
                        print(json.dumps({
                            "data": {
                                "title": "Quarterly Review Wiki",
                                "markdown": "# Intro\\nAlpha",
                                "has_more": True,
                                "next_offset": 1
                            }
                        }, ensure_ascii=False))
                    else:
                        print(json.dumps({
                            "data": {
                                "title": "Quarterly Review Wiki",
                                "markdown": "## Details\\nBeta",
                                "has_more": False
                            }
                        }, ensure_ascii=False))
                    sys.exit(0)

                if argv[:2] == ["slides", "+create"]:
                    print(json.dumps({
                        "data": {
                            "xml_presentation_id": "sldcnNew123",
                            "url": "https://example.feishu.cn/slides/sldcnNew123"
                        }
                    }, ensure_ascii=False))
                    sys.exit(0)

                if argv[:3] == ["slides", "xml_presentation.slide", "create"]:
                    state["slide_count"] += 1
                    state_path.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
                    print(json.dumps({
                        "data": {
                            "slide_id": f"slide_{state['slide_count']:03d}"
                        }
                    }, ensure_ascii=False))
                    sys.exit(0)

                print(json.dumps({"error": f"unexpected command: {argv}"}, ensure_ascii=False))
                sys.exit(1)
                """
            ),
            encoding="utf-8",
        )
        script_path.chmod(0o755)
        return bin_dir, log_path, state_path

    def run_script(self, workdir: Path, args: list[str], env: dict[str, str]) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(SCRIPT_PATH), *args],
            cwd=workdir,
            env=env,
            text=True,
            capture_output=True,
            check=False,
        )

    def test_cli_run_dir_semiflow_from_resolve_to_publish(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            bin_dir, log_path, state_path = self.write_fake_lark_cli(root)
            run_dir = root / "run"
            env = os.environ.copy()
            env["PATH"] = f"{bin_dir}{os.pathsep}{env['PATH']}"
            env["FAKE_LARK_CLI_LOG"] = str(log_path)
            env["FAKE_LARK_CLI_STATE"] = str(state_path)

            result = self.run_script(
                root,
                ["resolve-source", "--run-dir", str(run_dir), "--doc-name", "Quarterly Review"],
                env,
            )
            self.assertEqual("", result.stdout)
            self.assertEqual("", result.stderr)
            self.assertEqual(0, result.returncode)

            resolved_source = json.loads((run_dir / "resolved-source.json").read_text(encoding="utf-8"))
            self.assertTrue(resolved_source["needs_user_choice"])
            self.assertEqual(2, len(resolved_source["search_candidates"]))

            result = self.run_script(
                root,
                [
                    "choose-source",
                    "--resolved-source",
                    str(run_dir / "resolved-source.json"),
                    "--candidate-index",
                    "2",
                ],
                env,
            )
            self.assertEqual(0, result.returncode)
            resolved_source = json.loads((run_dir / "resolved-source.json").read_text(encoding="utf-8"))
            self.assertFalse(resolved_source["needs_user_choice"])
            self.assertEqual("wiki_url", resolved_source["resolved_kind"])
            self.assertEqual("Quarterly Review Wiki", resolved_source["title"])

            result = self.run_script(
                root,
                [
                    "fetch",
                    "--run-dir",
                    str(run_dir),
                    "--resolved-source",
                    str(run_dir / "resolved-source.json"),
                ],
                env,
            )
            self.assertEqual(0, result.returncode)
            source_md = (run_dir / "source.md").read_text(encoding="utf-8")
            self.assertEqual("# Intro\nAlpha\n\n## Details\nBeta\n", source_md)

            outline = {
                "presentation": {
                    "title": "Quarterly Review",
                    "target_mode": "new",
                    "content_mode": "report",
                    "max_slides": 2,
                    "source": {
                        "input_kind": resolved_source["input_kind"],
                        "resolved_kind": resolved_source["resolved_kind"],
                        "resolved_value": resolved_source["resolved_value"],
                    },
                },
                "slides": [
                    {
                        "no": 1,
                        "role": "cover",
                        "section_divider": False,
                        "title": "Quarterly Review",
                        "layout": "title-only",
                        "key_points": ["Alpha", "Beta"],
                    },
                    {
                        "no": 2,
                        "role": "content",
                        "section_divider": False,
                        "title": "Highlights",
                        "layout": "title-body",
                        "key_points": ["Alpha", "Beta"],
                    },
                ],
            }
            (run_dir / "outline.json").write_text(
                json.dumps(outline, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )

            result = self.run_script(
                root,
                ["validate-outline", "--outline", str(run_dir / "outline.json")],
                env,
            )
            self.assertEqual(0, result.returncode)

            result = self.run_script(
                root,
                ["render", "--outline", str(run_dir / "outline.json"), "--run-dir", str(run_dir)],
                env,
            )
            self.assertEqual(0, result.returncode)
            render_summary = json.loads((run_dir / "render-summary.json").read_text(encoding="utf-8"))
            self.assertEqual(2, render_summary["count"])
            self.assertEqual("briefing", render_summary["theme"])

            result = self.run_script(
                root,
                [
                    "publish",
                    "--outline",
                    str(run_dir / "outline.json"),
                    "--slides-json",
                    str(run_dir / "slides.json"),
                    "--run-dir",
                    str(run_dir),
                ],
                env,
            )
            self.assertEqual(0, result.returncode)
            publish_result = json.loads((run_dir / "publish-result.json").read_text(encoding="utf-8"))
            self.assertEqual("new", publish_result["target_mode"])
            self.assertEqual("sldcnNew123", publish_result["xml_presentation_id"])
            self.assertEqual(["slide_001", "slide_002"], publish_result["slide_ids"])

            command_log = [
                json.loads(line)
                for line in log_path.read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertEqual(
                [
                    "docs",
                    "+search",
                    "--as",
                    "user",
                    "--format",
                    "json",
                    "--query",
                    "Quarterly Review",
                ],
                command_log[0],
            )
            self.assertEqual(["wiki", "spaces", "get_node"], command_log[1][:3])
            self.assertEqual(["docs", "+fetch"], command_log[2][:2])
            self.assertEqual(["docs", "+fetch"], command_log[3][:2])
            self.assertEqual(
                [
                    "slides",
                    "+create",
                    "--as",
                    "user",
                    "--title",
                    "Quarterly Review",
                ],
                command_log[4],
            )
            self.assertEqual(["slides", "xml_presentation.slide", "create"], command_log[5][:3])
            self.assertEqual(["slides", "xml_presentation.slide", "create"], command_log[6][:3])


if __name__ == "__main__":
    unittest.main()
