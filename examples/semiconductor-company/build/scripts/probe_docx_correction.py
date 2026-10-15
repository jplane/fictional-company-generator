import json
import os
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(root.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts"))
import company
import playwright_capture

os.environ.update(OFFICECLI_SKIP_UPDATE="1", OFFICECLI_NO_AUTO_RESIDENT="1")
plan = company.load(root / "build" / "plan.json")
jobs = {j["document_id"]: j for j in company.load(root / "build" / "jobs-pilots.json")}
out = root / "build" / "docx-correction-probes"
out.mkdir(exist_ok=True)
with playwright_capture.CaptureEngine() as engine:
    for doc in plan["documents"]:
        if doc["id"] not in jobs or doc["format"] != "docx":
            continue
        path = out / f"{doc['id']}.docx"
        shutil.copy2(root / "documents" / doc["filename"], path)
        with zipfile.ZipFile(path) as package:
            tree = ET.fromstring(package.read("word/document.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = tree.findall("w:body/w:p", ns)
        commands = [{"command": "set", "path": f"/body/p[{i+1}]",
                     "props": {"firstLineIndent": "440"}} for i in range(len(paragraphs))]
        batch = out / f"{doc['id']}-indent.json"
        batch.write_text(json.dumps(commands), encoding="utf-8")
        for arguments in [
            ["batch", str(path), "--input", str(batch), "--stop-on-error"],
            ["validate", str(path)],
            ["view", str(path), "issues", "--limit", "100000"],
            ["view", str(path), "html", "--out", str(out / f"{doc['id']}.html")],
        ]:
            result = subprocess.run(["officecli", *arguments, "--json"], capture_output=True,
                                    text=True, check=True)
            if "html" in arguments:
                if not (out / f"{doc['id']}.html").is_file():
                    raise RuntimeError(result.stdout + result.stderr)
                continue
            parsed = json.loads(result.stdout)
            if not parsed.get("success"):
                raise RuntimeError(result.stdout)
            if "issues" in arguments and parsed["data"]["count"]:
                raise RuntimeError(result.stdout)
        result = engine.capture_docx_pages(out / f"{doc['id']}.html",
                                          out / doc["id"], jobs[doc["id"]]["render_targets"][0]["pages"])
        company.save(out / f"{doc['id']}-capture.json", result)
        print(doc["id"], "explicit indentation, issue scan and complete page capture passed")
