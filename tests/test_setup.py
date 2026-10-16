import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / ".github" / "skills" / "company-content-generator"


class SetupTests(unittest.TestCase):
    def test_portable_uv_project_has_optional_capture_dependencies(self):
        project = (SKILL / "pyproject.toml").read_text(encoding="utf-8")
        lock = (SKILL / "uv.lock").read_text(encoding="utf-8")
        for requirement in ('requires-python = ">=3.10"', "dependencies = []",
                            '[project.optional-dependencies]',
                            'capture = ["playwright==1.63.0"]', "package = false"):
            self.assertIn(requirement, project)
        self.assertIn('name = "playwright"\nversion = "1.63.0"', lock)
        self.assertNotIn("packagefeedproxy", project + lock)
        self.assertNotIn("visualstudio.com", project + lock)

    def test_claude_installation_includes_complete_skill(self):
        readme = " ".join((ROOT / "README.md").read_text(encoding="utf-8").split())
        for destination in ("<project>/.claude/skills/company-content-generator",
                            "~/.claude/skills/company-content-generator",
                            "<project>/.github/skills/company-content-generator",
                            "~/.copilot/skills/company-content-generator"):
            self.assertIn(destination, readme)
        for requirement in ("`references`", "`scripts`", "`pyproject.toml`",
                            "`uv.lock`", "/company-content-generator"):
            self.assertIn(requirement, readme)

    def test_execution_guidance_preserves_portable_storage(self):
        instructions = (SKILL / "SKILL.md").read_text(encoding="utf-8")
        guide = (SKILL / "references" / "dependencies.md").read_text(encoding="utf-8")
        for requirement in ("OS-specific separators", "uv run --locked --extra capture",
                            "relative filesystem paths always use `/`",
                            "existing stored evidence"):
            self.assertIn(requirement, instructions)
        for requirement in ("UV_PYTHON_DOWNLOADS=never", "UV_PROJECT_ENVIRONMENT",
                            "uv sync --locked --extra capture", "UV_DEFAULT_INDEX"):
            self.assertIn(requirement, guide)
        for old_instruction in ("Activate.ps1", "python3 -m venv", "python -m pip install",
                                "```powershell", "```sh\n"):
            self.assertNotIn(old_instruction, instructions + guide)

    def test_ci_uses_pinned_static_scanner_and_retains_failures(self):
        workflow = (ROOT / ".github" / "workflows" / "skillspector.yml").read_text(
            encoding="utf-8")
        for requirement in ("--no-llm --baseline .github/skillspector-baseline.json", "--locked --no-dev",
                            "--python 3.12", "--no-env-file", "if: always()",
                            "contents: read", "persist-credentials: false",
                            "LANGSMITH_TRACING: 'false'", "LANGCHAIN_TRACING_V2: 'false'",
                            "if-no-files-found: error", "baseline_sha256=",
                            "python .github/scripts/skillspector_summary.py"):
            self.assertIn(requirement, workflow)
        for forbidden in ("continue-on-error", "pull_request_target", "--fail-on-incomplete",
                          "--transitive", "secrets.", "|| true", "skillspector baseline"):
            self.assertNotIn(forbidden, workflow)
        self.assertRegex(workflow, r"ref: [0-9a-f]{40} # v")
        for action in re.findall(r"uses: (\S+)", workflow):
            self.assertRegex(action, r"@[0-9a-f]{40}$")

    def test_scanner_exceptions_are_exact_and_individually_reviewed(self):
        baseline = json.loads((ROOT / ".github" / "skillspector-baseline.json").read_text(
            encoding="utf-8"))
        self.assertEqual(baseline["version"], 2)
        self.assertEqual(baseline["scanner_version"], "2.12.0")
        self.assertEqual(baseline["rules"], [])
        entries = baseline["fingerprints"]
        self.assertEqual(len(entries), 23)
        self.assertEqual(len({entry["hash"] for entry in entries}), len(entries))
        for entry in entries:
            self.assertRegex(entry["hash"], r"^sha256:[0-9a-f]{64}$")
            self.assertTrue((SKILL / entry["file"]).is_file())
            self.assertGreater(len(entry["reason"]), 60)
            self.assertNotIn("auto-generated", entry["reason"])

    def test_scanner_summary_warns_on_incomplete_coverage_and_preserves_errors(self):
        for complete, successful in ((True, True), (False, True), (False, False)):
            with self.subTest(complete=complete, successful=successful):
                report = {
                    "risk_assessment": {"score": 0},
                    "suppressed_count": 23,
                    "execution_successful": successful,
                    "analysis_completeness": {
                        "coverage_percent": 100 if complete else 72.7,
                        "is_complete": complete,
                        "execution_successful": successful,
                        "ledger_exceptions": [] if complete else [
                            {"reason_code": "reference_missing"},
                            {"reason_code": "reference_missing"},
                            {"reason_code": "static_parse_limit"},
                        ],
                    },
                }
                with tempfile.TemporaryDirectory() as directory:
                    path = Path(directory) / "report.json"
                    summary = Path(directory) / "summary.md"
                    summary.write_text("Existing summary\n", encoding="utf-8")
                    path.write_text(json.dumps(report), encoding="utf-8")
                    env = dict(os.environ, GITHUB_STEP_SUMMARY=str(summary))
                    result = subprocess.run(
                        [sys.executable, str(ROOT / ".github" / "scripts" /
                                             "skillspector_summary.py"), str(path)],
                        env=env, capture_output=True, text=True, timeout=30)
                    self.assertEqual(result.returncode == 0, successful, result.stderr)
                    text = summary.read_text(encoding="utf-8")
                    self.assertTrue(text.startswith("Existing summary\n"))
                    self.assertIn("Reviewed findings suppressed: 23", text)
                    if not successful:
                        self.assertIn("execution failed", result.stderr)
                    elif not complete:
                        self.assertIn("::warning", result.stdout)
                        self.assertIn("INCOMPLETE; nonblocking", text)
                        self.assertIn("`reference_missing`: 2", text)
                        self.assertIn("`static_parse_limit`: 1", text)
                    else:
                        self.assertNotIn("::warning", result.stdout)
                        self.assertIn("100% (complete)", text)

    def test_scanner_summary_rejects_missing_or_malformed_reports(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            for content in (None, "{invalid", "{}"):
                with self.subTest(content=content):
                    if content is not None:
                        path.write_text(content, encoding="utf-8")
                    result = subprocess.run(
                        [sys.executable, str(ROOT / ".github" / "scripts" /
                                             "skillspector_summary.py"), str(path)],
                        env=dict(os.environ, GITHUB_STEP_SUMMARY=str(
                            Path(directory) / "summary.md")),
                        capture_output=True, text=True, timeout=30)
                    self.assertNotEqual(result.returncode, 0)

    @unittest.skipUnless(shutil.which("uv"), "uv is required for copied-skill execution")
    def test_copied_skill_runs_from_unrelated_directory_without_capture(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            copied = root / "project with spaces" / ".claude" / "skills" / SKILL.name
            shutil.copytree(SKILL, copied, ignore=shutil.ignore_patterns(
                ".venv", "__pycache__", "*.pyc"))
            caller = root / "unrelated working directory"
            caller.mkdir()
            env = dict(os.environ, UV_PYTHON_DOWNLOADS="never",
                       UV_PROJECT_ENVIRONMENT=str(root / "managed environment"),
                       PYTHONDONTWRITEBYTECODE="1")
            for variable in ("UV_DEFAULT_INDEX", "UV_INDEX_URL", "UV_INDEX", "VIRTUAL_ENV"):
                env.pop(variable, None)
            prefix = [shutil.which("uv"), "run", "--locked", "--offline",
                      "--project", str(copied), "--python", sys.executable]
            for script in ("company.py", "officecli_runner.py"):
                result = subprocess.run(
                    [*prefix, str(copied / "scripts" / script), "--help"],
                    cwd=caller, env=env, capture_output=True, text=True, timeout=60)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIn("usage:", result.stdout)
            result = subprocess.run(
                [*prefix, "python", "-c",
                 "import importlib.util; assert importlib.util.find_spec('playwright') is None"],
                cwd=caller, env=env, capture_output=True, text=True, timeout=60)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertFalse((caller / ".venv").exists())


if __name__ == "__main__":
    unittest.main()
