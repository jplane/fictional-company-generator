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
        for requirement in ("--no-llm --fail-on-incomplete", "--locked --no-dev",
                            "--python 3.12", "--no-env-file", "if: always()",
                            "contents: read", "persist-credentials: false",
                            "LANGSMITH_TRACING: 'false'", "LANGCHAIN_TRACING_V2: 'false'",
                            "if-no-files-found: error"):
            self.assertIn(requirement, workflow)
        for forbidden in ("continue-on-error", "pull_request_target", "--baseline",
                          "--transitive", "secrets.", "|| true"):
            self.assertNotIn(forbidden, workflow)
        self.assertRegex(workflow, r"ref: [0-9a-f]{40} # v")
        for action in re.findall(r"uses: (\S+)", workflow):
            self.assertRegex(action, r"@[0-9a-f]{40}$")

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
