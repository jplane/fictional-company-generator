"""Checked OfficeCLI execution, capability probes, and resumable corpus mechanics."""

import argparse
import copy
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from contextlib import contextmanager
from contextvars import copy_context
from datetime import datetime, timezone
from functools import wraps
from pathlib import Path
from threading import Lock
from uuid import uuid4
from xml.etree import ElementTree

import company
import playwright_capture


def timestamp():
    return datetime.now(timezone.utc).isoformat()


def write_json(path, value):
    """Atomically persist mutable runner-owned state; canon/plan still use no-overwrite save."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", newline="\n",
                                         dir=path.parent, suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(value, stream, indent=2, ensure_ascii=True, allow_nan=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


@contextmanager
def run_lock(root):
    path = Path(root).resolve() / "build" / "runner.lock"
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+b") as stream:
        stream.seek(0, os.SEEK_END)
        if not stream.tell():
            stream.write(b"1")
            stream.flush()
        stream.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise company.Invalid("Run is owned by another invocation; retry after it exits") from exc
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == "nt":
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream, fcntl.LOCK_UN)


def phase_event(root, action, name=None, event_id=None, parent=None, pass_name="generation",
                document_id=None, status="passed"):
    path = Path(root).resolve() / "build" / "timings.json"
    company.require((Path(root) / "canon.json").is_file(), "Initialize the corpus before recording phases")
    with run_lock(root):
        receipt = company.load(path) if path.exists() else {"timing_version": 1, "events": []}
        if action == "start":
            company.require(company.text(name), "Phase name required")
            company.require(parent is None or any(e["id"] == parent and e["status"] == "running"
                                                 for e in receipt["events"]), "Unknown/inactive parent phase")
            event = {"id": str(uuid4()), "name": name, "parent": parent, "pass": pass_name,
                     "document_id": document_id, "started_at": timestamp(), "status": "running"}
            receipt["events"].append(event)
        else:
            matches = [e for e in receipt["events"] if e["id"] == event_id]
            company.require(len(matches) == 1 and matches[0]["status"] == "running",
                            "Unknown/already ended phase")
            event = matches[0]
            company.require(not any(e.get("parent") == event_id and e["status"] == "running"
                                    for e in receipt["events"]), "End child phases before their parent")
            event.update(ended_at=timestamp(), status=status)
            event["elapsed_seconds"] = (datetime.fromisoformat(event["ended_at"])
                                        - datetime.fromisoformat(event["started_at"])).total_seconds()
        write_json(path, receipt)
        return event


def timing_summary(root):
    build = Path(root).resolve() / "build"
    state = company.load(build / "runner-state.json") if (build / "runner-state.json").exists() else {}
    receipt = company.load(build / "timings.json") if (build / "timings.json").exists() else {"events": []}
    events = receipt["events"]
    parents = {e.get("parent") for e in events}
    operations = [op for doc in state.get("documents", {}).values()
                  for op in doc.get("operations", []) + [
                      e for e in doc.get("qa_state", {}).get("operations", []) if e.get("pass") == "qa"]]
    captures = [op for doc in state.get("documents", {}).values()
                for op in doc.get("captures", []) + [
                    e for e in doc.get("qa_state", {}).get("captures", []) if e.get("pass") == "qa"]]
    by_operation = {}
    for op in operations:
        entry = by_operation.setdefault(op["operation"], {"attempts": 0, "worker_active_seconds": 0})
        entry["attempts"] += 1
        entry["worker_active_seconds"] += op.get("active_seconds", 0)
    return {"timing_version": 1,
            "runner_active_wall_seconds": state.get("active_wall_seconds"),
            "between_invocations_seconds": state.get("between_invocations_seconds", 0),
            "subprocess_worker_active_seconds": sum(e.get("active_seconds", 0) for e in operations),
            "subprocess_operations": by_operation,
            "capture_elapsed_worker_seconds": sum(e.get("elapsed_seconds", e.get("active_seconds", 0))
                                                   for e in captures),
            "capture_worker_active_seconds": sum(e.get("lane_timing", {}).get("active_seconds", 0)
                                                  for e in captures),
            "capture_queue_seconds": sum(e.get("lane_timing", {}).get("waiting_seconds", 0) for e in captures),
            "unseparated_capture_attempts": sum("lane_timing" not in e for e in captures),
            "captured_images": sum(e.get("image_count", 0) for e in captures),
            "inspection_cache_hits": sum(d.get("inspection_cache_hits", 0)
                                         for d in state.get("documents", {}).values()),
            "validation_counts": state.get("validation_counts", {}),
            "runner_measurements": state.get("measurements", []),
            "document_spans": state.get("phases", []),
            "leaf_agent_spans": [e for e in events if e["id"] not in parents],
            "unknown_agent_phases": not bool(events),
            "milestones": {name: company.load(build / filename) for name, filename in (
                ("generation", "generation.json"), ("qa", "qa-state.json")) if (build / filename).exists()},
            "limitations": "Agent spans are elapsed observations, not inferred active time. "
                           "Worker totals may overlap; never add them to wall time or parent spans."}


def recorded_phase(name):
    def decorate(function):
        @wraps(function)
        def invoke(self, *args, **kwargs):
            self.boundary()
            event = {"id": str(uuid4()), "phase": name, "pass": "qa" if self.qa_mode else "generation",
                     "document_id": args[0].get("document_id") if args and isinstance(args[0], dict) else None,
                     "started_at": timestamp(), "status": "running", "active_seconds": 0}
            if getattr(self, "worker_id", None):
                event.update(family=list(self.families[self.worker_id]),
                             queued_at=self.queued_at,
                             queue_seconds=max(0, (datetime.fromisoformat(event["started_at"])
                                                   - datetime.fromisoformat(self.queued_at)).total_seconds()))
            self.state.setdefault("phases", []).append(event)
            self.active_phase = (event, time.monotonic())
            self.checkpoint()
            try:
                result = function(self, *args, **kwargs)
                event["status"] = "passed"
                return result
            finally:
                event.update(ended_at=timestamp(), active_seconds=time.monotonic() - self.active_phase[1])
                if event["status"] == "running":
                    event["status"] = "blocked"
                self.active_phase = None
                self.checkpoint()
        return invoke
    return decorate


def setup_action(installed, available, compatible, requested_update=False, authorized=False):
    """Make setup decisions explicitly; acquisition is separate from corpus execution."""
    needed = not installed or not compatible or requested_update
    if needed:
        company.require(available, "Required OfficeCLI install/update has no verified available release")
        return "install" if authorized and not installed else (
            "update" if authorized else "authorization_required")
    return "offer_update" if available and available != installed else (
        "update_lookup_unavailable" if available is None else "use_installed")


def verify_release(asset, expected_sha256, observed_version, expected_version):
    company.require(isinstance(expected_sha256, str)
                    and len(expected_sha256) == 64
                    and company.file_hash(asset) == expected_sha256.lower(),
                    "OfficeCLI release checksum missing/mismatched")
    company.require(observed_version.strip() == expected_version,
                    "OfficeCLI installed version mismatch; preserve prior binary")


def response(stdout, operation):
    return company.validate_officecli_result(stdout, operation)


class OfficeCLI:
    def __init__(self, executable=None, timeout=120):
        resolved = shutil.which(executable or "officecli")
        company.require(resolved, "OfficeCLI unavailable; complete authorized dependency setup first")
        self.executable = str(Path(resolved).resolve())
        company.require(timeout > 0, "OfficeCLI timeout must be positive")
        self.timeout = timeout
        self.env = dict(os.environ, OFFICECLI_SKIP_UPDATE="1", OFFICECLI_NO_AUTO_RESIDENT="1")
        version = subprocess.run([self.executable, "--version"], capture_output=True, text=True,
                                 encoding="utf-8", env=self.env, timeout=timeout, check=True)
        self.tool = {"version": version.stdout.strip(), "sha256": company.file_hash(self.executable)}

    def run(self, arguments, operation):
        company.require(company.file_hash(self.executable) == self.tool["sha256"],
                        "OfficeCLI executable drift; revalidate capabilities before resuming")
        forbidden = {"refresh", "--page-count", "--best-effort", "--force", "--browser",
                    "--render", "screenshot"}
        company.require(not forbidden.intersection(arguments),
                        "Noncompliant OfficeCLI operation: pixel capture is a Playwright-adapter "
                        "responsibility, not an OfficeCLI screenshot/render/auto-detect call")
        if "html" in arguments:
            company.require("--out" in arguments, "HTML export must write a bound output file, not stdout")
        started = timestamp()
        begin = time.monotonic()
        result = subprocess.run([self.executable, *arguments, "--json"], capture_output=True,
                                text=True, encoding="utf-8", env=self.env, timeout=self.timeout)
        stdout = result.stdout
        if result.returncode == 0 and "html" in arguments:
            # Observed OfficeCLI behavior: `view <file> html --out <path>` prints the bound
            # output path as plain text on success (even with --json), unlike every other
            # mode. Normalize to the standard {"success","data"} envelope at this one boundary
            # so downstream validation never special-cases a non-JSON OfficeCLI response.
            try:
                json.loads(stdout)
            except json.JSONDecodeError:
                stdout = json.dumps({"success": True, "data": stdout.strip()})
        log = {"arguments": arguments, "operation": operation, "started_at": started,
               "ended_at": timestamp(), "active_seconds": time.monotonic() - begin,
               "waiting_seconds": 0, "returncode": result.returncode,
               "stdout": stdout, "stderr": result.stderr, "tool": self.tool}
        # Return failed raw evidence too; callers persist it before interpreting success.
        return log


def checked(log):
    company.require(log["returncode"] == 0 and not log["stderr"].strip(),
                    f"{log['operation']}: OfficeCLI failed: {log['stderr'] or log['stdout']}")
    return response(log["stdout"], log["operation"])


def formula_values(path, bindings, claims=None):
    return company.verify_formula_values(path, bindings, claims)


def no_page_fields(path):
    if path.suffix.lower() != ".docx":
        return
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith("word/") and name.endswith(".xml"):
                element = ElementTree.fromstring(package.read(name))
                company.require(not element.findall(".//w:fldChar", company.NS)
                                and not element.findall(".//w:fldSimple", company.NS),
                                "DOCX fields require an explicitly proven Office-free capability; refresh is forbidden")


def fixture_commands(fmt):
    if fmt == "docx":
        return [{"command": "add", "parent": "/body", "type": "paragraph",
                 "props": {"text": "OfficeCLI capability probe: long title and -2.148",
                           "font": "Arial", "size": "16pt", "bold": "true", "firstLineIndent": "0.5cm"}}]
    if fmt == "pptx":
        return [{"command": "add", "parent": "/", "type": "slide", "props": {"title": "Capability probe"}},
                {"command": "add", "parent": "/slide[1]", "type": "shape",
                 "props": {"text": "Long title and -2.148", "x": "2cm", "y": "4cm", "size": "20pt"}}]
    return [{"command": "add", "parent": "/", "type": "sheet", "props": {"name": "Results"}},
            {"command": "set", "path": "/Sheet1/A1", "props": {"value": "-2.148"}},
            {"command": "set", "path": "/Sheet1/A2", "props": {"formula": "A1*2"}},
            {"command": "set", "path": "/Results/A1", "props": {"formula": "Sheet1!A2*2"}},
            # Minimal real column + line two-cell-anchor chart objects: these are the only
            # supported chart capability tested/certified by this preflight (see
            # chart_capability_probe below). Placed on Sheet1 so they never collide with the
            # basic region probe, which captures Results!A1 on a different sheet.
            {"command": "set", "path": "/Sheet1/C1", "props": {"value": "Q1"}},
            {"command": "set", "path": "/Sheet1/D1", "props": {"value": "10"}},
            {"command": "set", "path": "/Sheet1/C2", "props": {"value": "Q2"}},
            {"command": "set", "path": "/Sheet1/D2", "props": {"value": "20"}},
            {"command": "add", "parent": "/Sheet1", "type": "chart",
             "props": {"dataRange": "Sheet1!D1:D2", "categories": "Sheet1!C1:C2",
                       "chartType": "column", "anchor": "F2:L18", "title": "Capability Probe Column"}},
            {"command": "add", "parent": "/Sheet1", "type": "chart",
             "props": {"dataRange": "Sheet1!D1:D2", "categories": "Sheet1!C1:C2",
                       "chartType": "line", "anchor": "F20:L36", "title": "Capability Probe Line"}}]


def html_export_arguments(path, out_html):
    return ["view", str(path), "html", "--out", str(out_html)]


def capture_probe(engine, html_path, fmt, image, tmp_dir):
    """One bound Playwright capture per capability probe; never a fabricated
    OfficeCLI subprocess result. Writes the final probe image at `image`."""
    source_html_sha256 = company.file_hash(html_path)
    started = timestamp()
    begin = time.monotonic()
    if fmt == "docx":
        result = engine.capture_docx_pages(html_path, tmp_dir, expected_pages=1)
        shutil.move(result["pages"][0]["path"], image)
        unit, rendered_count = "pages", result["rendered_pages"]
    elif fmt == "pptx":
        result = engine.capture_pptx_slides(html_path, tmp_dir, expected_slides=1)
        shutil.move(result["slides"][0]["path"], image)
        unit, rendered_count = "slides", result["rendered_slides"]
    else:
        region = engine.capture_xlsx_region(html_path, image, "Results", "A1", "A1", sheet_index=1)
        unit, rendered_count = "region", None
    log = {"operation": "capture", "status": "passed", "started_at": started, "ended_at": timestamp(),
           "active_seconds": time.monotonic() - begin, "tool": engine.identity,
           "source_html_sha256": source_html_sha256, "unit": unit,
           "image_sha256": company.file_hash(image)}
    if rendered_count is not None:
        log["rendered_count"] = rendered_count
    return log


def chart_capability_probe(engine, path, html_path, source_html_sha256, out_dir):
    """Real column+line two-cell-anchor chart capture proof for the preflight's xlsx
    probe fixture (see fixture_commands). A basic cell-only region capture does not
    exercise the chart-object capture path at all, so it must never be treated as
    chart capability evidence; this is the one place that actually binds it."""
    with zipfile.ZipFile(path) as package:
        inventory = company.chart_object_inventory(package)
    supported = {o["chart_type"]: o for o in inventory if o["status"] == "supported"}
    company.require({"column", "line"} <= set(supported),
                    "Preflight chart fixture did not yield both a supported column and line chart object; "
                    "chart capability cannot be certified")
    out_dir = Path(out_dir)
    entries = []
    for chart_type in ("column", "line"):
        obj = supported[chart_type]
        target = {"id": f"chart-probe-{chart_type}", "type": "chart",
                  "sheet": obj["sheet"], "shape_id": obj["shape_id"]}
        chart_inventory = {(obj["sheet"], obj["shape_id"]): obj}
        images, capture_log = capture_target(engine, html_path, "xlsx", target,
                                             out_dir / chart_type, source_html_sha256, chart_inventory)
        company.validate_capture_result(capture_log)
        company.require(capture_log.get("clipped") is False and capture_log.get("occluded") is False,
                        f"Preflight {chart_type} chart capability probe is clipped/occluded; "
                        "chart capability cannot be certified from a degraded capture")
        entries.append({"chart_type": chart_type, "sheet": obj["sheet"], "shape_id": obj["shape_id"],
                        "dom_path": capture_log["dom_path"], "image_sha256": images[0]["sha256"],
                        "bounds": capture_log["bounds"], "clipped": capture_log["clipped"],
                        "occluded": capture_log["occluded"], "tool": capture_log["tool"]})
    return entries


def preflight(cli, formats, out):
    company.require(formats and set(formats) <= set(company.FORMATS), "Invalid preflight formats")
    out = Path(out)
    company.require(not out.exists(), "Preflight already exists; use a new capability revision")
    receipt = {"preflight_version": 2, "engine": "OfficeCLI", "tool": cli.tool,
               "formats": sorted(set(formats)), "status": "pending", "started_at": timestamp(),
               "source": "https://github.com/iOfficeAI/OfficeCLI", "probes": []}
    company.save(out, receipt)
    engine = None
    try:
        with tempfile.TemporaryDirectory(prefix="officecli-probe-", dir=out.parent) as temporary:
            for fmt in sorted(set(formats)):
                path = Path(temporary) / f"probe.{fmt}"
                html_path = Path(temporary) / f"probe-{fmt}.html"
                image = out.parent / f"{out.stem}-{fmt}.png"
                company.require(not image.exists(), "Preflight image already exists; use a new revision")
                probe = {"format": fmt, "logs": [], "formula_count": 0, "image": image.name,
                         "features": ["basic", "html_screenshots"] +
                                     (["ordered_cross_sheet_formulas"] if fmt == "xlsx" else [])}
                receipt["probes"].append(probe)

                def invoke(arguments, operation):
                    result = cli.run(arguments, operation)
                    probe["logs"].append(result)
                    write_json(out, receipt)
                    checked(result)

                invoke(["create", str(path)], "create")
                invoke(["batch", str(path), "--commands", json.dumps(fixture_commands(fmt)),
                        "--stop-on-error"], "author")
                for operation, arguments in (("validate", ["validate", str(path)]),
                                              ("issues", ["view", str(path), "issues"]),
                                              ("extract", ["view", str(path), "text"])):
                    invoke(arguments, operation)
                if fmt == "xlsx":
                    probe["formula_count"] = formula_values(path, [
                        {"sheet": "Sheet1", "cell": "A2", "formula": "A1*2", "value": "-4.296"},
                        {"sheet": "Results", "cell": "A1", "formula": "Sheet1!A2*2", "value": "-8.592"}])["formula_count"]
                    invoke(["view", str(path), "issues", "--type", "chart_cache_stale"], "chart_cache_stale")
                invoke(html_export_arguments(path, html_path), "html_export")
                if engine is None:
                    # One managed Chromium launch, reused across every remaining probe/render.
                    engine = playwright_capture.CaptureEngine()
                    receipt["capture_tool"] = engine.identity
                    write_json(out, receipt)
                capture_log = capture_probe(engine, html_path, fmt, image, Path(temporary) / f"capture-{fmt}")
                probe["capture"] = capture_log
                company.require(image.is_file() and image.read_bytes().startswith(b"\x89PNG\r\n\x1a\n"),
                                "Playwright did not produce a readable PNG; preflight is blocked")
                company.validate_capture_result(capture_log)
                probe["image_sha256"] = company.file_hash(image)
                if fmt == "xlsx":
                    # A basic cell-only region probe above never exercises chart-object
                    # capture; certify "chart" only from this dedicated, bound evidence.
                    chart_entries = chart_capability_probe(
                        engine, path, html_path, capture_log["source_html_sha256"],
                        Path(temporary) / "capture-xlsx-chart")
                    for entry in chart_entries:
                        permanent = out.parent / f"{out.stem}-xlsx-chart-{entry['chart_type']}.png"
                        company.require(not permanent.exists(),
                                        "Preflight chart image already exists; use a new revision")
                        shutil.move(str(Path(temporary) / "capture-xlsx-chart" /
                                        entry["chart_type"] / "chart.png"), permanent)
                        entry["image"] = permanent.name
                        entry["image_sha256"] = company.file_hash(permanent)
                    probe["chart_capture"] = chart_entries
                    probe["features"].append("chart")
                write_json(out, receipt)
    except (company.Invalid, OSError, subprocess.SubprocessError,
            playwright_capture.CaptureError) as exc:
        receipt.update(status="blocked", blocker=str(exc), ended_at=timestamp())
        write_json(out, receipt)
        raise
    finally:
        if engine is not None:
            engine.close()
    receipt.update(status="passed", ended_at=timestamp())
    write_json(out, receipt)
    return receipt


def local_input(root, relative, *, native_cli=False):
    company.require(company.text(relative), "Missing runner input path")
    path = Path(relative).resolve() if native_cli and Path(relative).is_absolute() else company.stored_path(root, relative)
    company.require_build_path(path, root)
    company.require(path.is_file(), f"Missing runner input: {relative}")
    return path


def job_identity(root, job, canon, plan, modern=False):
    source_keys = ("batch_file", "template", "data_file", "layout_file")
    sources = {key: company.file_hash(local_input(root, job[key])) for key in source_keys if key in job}
    company.require("layout_file" in sources, "Every job requires its frozen layout file")
    return company.fingerprint({"canon": company.fingerprint(canon), "plan": company.fingerprint(plan),
                                "draft": company.file_hash(Path(root) / "build" / "drafts" /
                                                          f"{job['document_id']}.json"),
                                "job": {k: v for k, v in job.items()
                                        if k not in ({"correction_file", "correction_reason", "retry_reasons"}
                                                     | ({"export_bindings", "occurrences"} if modern else set()))},
                                "sources": sources})


def batch_file(root, relative):
    path = local_input(root, relative)
    commands = company.load(path)
    company.require(isinstance(commands, list) and commands, "Batch must be a nonempty JSON array")
    for command in commands:
        company.require(isinstance(command, dict)
                        and command.get("command") in ("add", "set", "remove", "move", "swap"),
                        "Only checked DOM mutations are supported; no raw cache injection or nested execution")
    return path


def capture_target(engine, html_path, fmt, target, out_dir, source_html_sha256, chart_inventory=None):
    company.require(company.text(target.get("id")), "Render target requires a stable ID")
    started = timestamp()
    begin = time.monotonic()
    lane_timing = None
    if fmt == "docx":
        result = engine.capture_docx_pages(html_path, out_dir, expected_pages=target["pages"])
        lane_timing = result.get("lane_timing")
        images = [{"page": p["page"], "path": p["path"], "sha256": p["sha256"], "overflow": p["overflow"]}
                  for p in result["pages"]]
        unit, rendered_count, extra = "pages", result["rendered_pages"], {}
    elif fmt == "pptx":
        result = engine.capture_pptx_slides(html_path, out_dir, expected_slides=target["slides"])
        lane_timing = result.get("lane_timing")
        images = [{"slide": s["slide"], "path": s["path"], "sha256": s["sha256"]} for s in result["slides"]]
        unit, rendered_count, extra = "slides", result["rendered_slides"], {}
    elif fmt == "xlsx" and target.get("type") == "chart":
        sheet, shape_id = target.get("sheet"), target.get("shape_id")
        company.require(company.text(sheet) and company.text(shape_id),
                        "XLSX chart render target needs a source sheet and shape_id")
        key = (sheet, shape_id)
        company.require(isinstance(chart_inventory, dict) and key in chart_inventory,
                        f"Chart render target {key} missing from the freshly verified chart inventory; "
                        "re-run coverage verification before capture")
        dom_path = chart_inventory[key]["expected_dom_path"]
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_png = out_dir / "chart.png"
        captured = engine.capture_xlsx_chart(html_path, out_png, sheet, dom_path, shape_id,
                                             sheet_index=target.get("sheet_index"))
        lane_timing = captured.get("lane_timing")
        images = [{"path": captured["path"], "sha256": captured["sha256"]}]
        unit, rendered_count = "chart", None
        extra = {"sheet": sheet, "shape_id": shape_id, "dom_path": captured["dom_path"],
                "bounds": captured["bounds"], "clipped": captured["clipped"],
                "occluded": captured["occluded"]}
    else:
        company.require(fmt == "xlsx" and target.get("type", "range") == "range",
                        "Unknown render target type")
        region = target.get("range")
        company.require(isinstance(region, str) and "!" in region, "XLSX render target needs Sheet!A1:C3")
        sheet, cells = region.rsplit("!", 1)
        addresses = cells.split(":")
        company.require(len(addresses) == 2, "XLSX render range must have two addresses")
        start_cell, end_cell = addresses
        start, end = map(company.cell_coordinates, addresses)
        company.require(start[0] <= end[0] and start[1] <= end[1], "Reversed render range")
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_png = out_dir / "region.png"
        captured = engine.capture_xlsx_region(html_path, out_png, sheet, start_cell, end_cell,
                                              sheet_index=target.get("sheet_index"))
        lane_timing = captured.get("lane_timing")
        images = [{"path": captured["path"], "sha256": captured["sha256"]}]
        unit, rendered_count = "region", None
        extra = {"start_cell": start_cell, "end_cell": end_cell}
        if "clip" in captured:
            extra["clip"] = captured["clip"]
        if "viewport" in captured:
            extra["viewport"] = captured["viewport"]
    log = {"operation": "capture", "status": "passed", "started_at": started, "ended_at": timestamp(),
           "active_seconds": time.monotonic() - begin, "tool": engine.identity,
           "source_html_sha256": source_html_sha256, "unit": unit, "image_sha256": images[0]["sha256"]}
    if rendered_count is not None:
        log["rendered_count"] = rendered_count
    if lane_timing is not None:
        log["lane_timing"] = lane_timing
        log.update(elapsed_seconds=log["active_seconds"], active_seconds=lane_timing["active_seconds"],
                   waiting_seconds=lane_timing["waiting_seconds"])
    log.update(extra)
    return images, log


def verify_regions(path, targets):
    return company.verify_render_targets(path, targets)


class Runner:
    def __init__(self, root, cli):
        self.root, self.cli = Path(root).resolve(), cli
        self.build = self.root / "build"
        self.canon = company.load(self.root / "canon.json")
        self.plan = company.load(self.build / "plan.json")
        self.docs = company.validate_plan(self.canon, self.plan)
        self.execution = company.load(self.build / "execution-plan.json")
        self.modern = self.execution.get("execution_version") == company.GENERATION_EXECUTION_VERSION
        self.qa_mode = False
        self.families = company.validate_execution(self.canon, self.plan, self.execution)
        self.base_selection = {e["document_id"]: e for e in company.review_selection(
            self.canon, self.plan, self.build / "drafts", self.execution)}
        self.capability = company.load(self.build / "preflight.json")
        company.require(self.capability.get("status") == "passed"
                        and self.capability.get("engine") == "OfficeCLI"
                        and self.capability.get("tool") == cli.tool
                        and {d["format"] for d in self.docs.values()} <= set(self.capability.get("formats", [])),
                        "OfficeCLI capability receipt absent/stale; rerun preflight")
        company.validate_preflight(self.capability, self.root)
        company.validate_required_features(self.execution, self.capability)
        self.state_path = self.build / "runner-state.json"
        self.state = company.load(self.state_path) if self.state_path.exists() else {
            "runner_version": 2 if self.modern else 1,
            **company.receipt_header(self.canon, self.plan), "documents": {}}
        company.require(self.state.get("runner_version") == (2 if self.modern else 1)
                        and all(self.state.get(k) == v for k, v in
                        company.receipt_header(self.canon, self.plan).items()), "Stale runner canon/plan")
        if self.modern:
            identity = company.generation_identity(self.execution)
            company.require(self.state.get("generation_sha256") in (None, identity),
                            "Changed generation contract; preserve outputs and create an affected revision")
            self.state["generation_sha256"] = identity
        execution_hash = company.fingerprint(self.execution)
        if self.state.get("execution_sha256") not in (None, execution_hash):
            company.require(self.execution["revision"] > self.state["execution_revision"],
                            "Changed execution policy/layout needs an explicit newer revision")
        self.state.update(execution_sha256=execution_hash, execution_revision=self.execution["revision"])
        self.started = time.monotonic()
        self.active_phase = None
        for event in self.state.get("phases", []):
            if event.get("status") == "running":
                event["status"] = "interrupted"
        capability_path = self.build / "capabilities" / f"{company.fingerprint(self.capability)}.json"
        if capability_path.exists():
            company.require(company.load(capability_path) == self.capability, "Changed pinned capability snapshot")
        else:
            company.save(capability_path, self.capability)
        self.capability_ref = {"path": capability_path.relative_to(self.root).as_posix(),
                               "sha256": company.file_hash(capability_path)}
        if self.state.get("last_checkpoint_at"):
            wait = (datetime.now(timezone.utc) -
                    datetime.fromisoformat(self.state["last_checkpoint_at"])).total_seconds()
            self.state["between_invocations_seconds"] = self.state.get("between_invocations_seconds", 0) + max(0, wait)
        self._engine = None

    def engine(self):
        """One managed Chromium browser, lazily created and reused across every
        render in this run; closed explicitly by the caller via close()."""
        if self._engine is None:
            self._engine = playwright_capture.CaptureEngine()
        return self._engine

    def close(self):
        if self._engine is not None:
            self._engine.close()
            self._engine = None

    def checkpoint(self):
        if getattr(self, "active_phase", None):
            self.active_phase[0]["active_seconds"] = time.monotonic() - self.active_phase[1]
        self.state["last_checkpoint_at"] = timestamp()
        write_json(self.state_path, self.state)

    def boundary(self):
        budget = self.execution.get("budget_seconds")
        if "phases" in self.state:
            elapsed = sum(event["active_seconds"] for event in self.state["phases"])
            if getattr(self, "active_phase", None):
                elapsed += time.monotonic() - self.active_phase[1] - self.active_phase[0]["active_seconds"]
        else:
            elapsed = sum(stage.get("active_seconds", 0) for state in self.state["documents"].values()
                          for stage in state.get("operations", []))
        company.require(budget is None or elapsed < budget,
                        "Runtime budget reached at a safe checkpoint; run remains partial")

    def operation(self, state, arguments, operation, corrective=None, repeat_success=False,
                  artifact_sha256=None):
        attempts = state.setdefault("attempts", {}).get(operation, 0)
        history = [entry for entry in state.get("operations", [])
                   if entry.get("operation") == operation]
        budget_attempts = (max(0, attempts - sum(entry.get("status") == "passed" for entry in history))
                           if repeat_success else attempts)
        company.require(budget_attempts < 2, f"{operation}: corrective retry already consumed")
        retry = bool(budget_attempts) and (not repeat_success or not history
                                         or history[-1].get("status") != "passed")
        if retry:
            reason = state.get("retry_reasons", {}).get(operation)
            company.require(corrective or company.text(reason),
                            f"{operation}: initial attempt failed; supply a concrete correction")
            arguments = corrective or arguments
        state["attempts"][operation] = attempts + 1
        self.checkpoint()
        try:
            log = self.cli.run(arguments, operation)
        except subprocess.TimeoutExpired as exc:
            state.setdefault("operations", []).append({"operation": operation, "status": "timeout",
                "ended_at": timestamp(), "active_seconds": self.cli.timeout,
                "reason": str(exc), "corrective": retry,
                "artifact_sha256": artifact_sha256, "repeat_success": repeat_success})
            self.checkpoint()
            raise company.Invalid(f"{operation}: timed out; checkpoint preserved") from exc
        log["corrective"] = retry
        log["pass"] = "qa" if getattr(self, "qa_mode", False) else "generation"
        if repeat_success:
            log.update(repeat_success=True, artifact_sha256=artifact_sha256)
        if retry:
            log["retry_reason"] = state.get("retry_reasons", {}).get(operation) or state.get("correction_reason")
        state.setdefault("operations", []).append(log)
        self.checkpoint()
        try:
            checked(log)
        except company.Invalid as exc:
            log.update(status="blocked", blocker=str(exc))
            self.checkpoint()
            raise
        log["status"] = "passed"
        self.checkpoint()
        return log

    def inspection(self, state, arguments, operation, digest):
        previous = [entry for entry in state["operations"] if entry.get("operation") == operation
                    and entry.get("status") == "passed" and entry.get("artifact_sha256") == digest
                    and entry.get("tool") == self.cli.tool]
        if previous:
            state["inspection_cache_hits"] = state.get("inspection_cache_hits", 0) + 1
            checked(previous[-1])
            return previous[-1]
        self.boundary()
        log = self.operation(state, arguments, operation, repeat_success=True,
                             artifact_sha256=digest)
        log.update(status="passed", artifact_sha256=digest,
                   evidence="Checked complete OfficeCLI result without errors/caveats")
        self.checkpoint()
        return log

    def formulas(self, state, path, job, draft):
        claims = {c["id"]: c for c in draft["claims"]}
        bindings = job.get("expected_formulas", [])
        try:
            result = formula_values(path, bindings, claims)
        except company.Invalid as exc:
            company.require(not getattr(self, "qa_mode", False),
                            f"QA cannot modify saved formulas: {exc}; create an affected revision")
            company.require(any(fragment in str(exc) for fragment in (
                "formula has no saved cached value", "stale/incorrect formula cache",
                "missing/error/unsupported formula cache")), f"Formula validation blocked: {exc}")
            company.require(not state.get("formula_retry"), "Formula corrective retry already consumed")
            company.require(bindings, "Missing grounded formula bindings; no corrective rewrite")
            # Expected bindings were validated before reading the defective saved caches.
            state["formula_retry"] = {"reason": str(exc), "started_at": timestamp(), "status": "pending"}
            self.checkpoint()
            commands = [{"command": "set", "path": f"/{b['sheet']}/{b['cell']}",
                         "props": {"formula": b["formula"].lstrip("=")}} for b in bindings]
            self.operation(state, ["batch", str(path), "--commands", json.dumps(commands),
                                  "--stop-on-error"], "formula_correction")
            state["artifact_sha256"] = company.file_hash(path)
            self.checkpoint()
            result = formula_values(path, bindings, claims)
            state["formula_retry"].update(status="passed", ended_at=timestamp())
            self.checkpoint()
        result["bindings"] = bindings
        return result

    def selection_for(self, identifier):
        selected = dict(self.base_selection[identifier])
        affected = {self.families[i] for i, state in self.state["documents"].items()
                    if self.corrected_failure(state)}
        if self.families[identifier] in affected:
            selected.update(content_required=True, visual_required=True,
                            reasons=sorted(set(selected["reasons"]) | {"affected_family"}))
        return selected

    @staticmethod
    def corrected_failure(state):
        repeatable = {entry["operation"] for entry in state.get("operations", [])
                      if entry.get("repeat_success")}
        return bool(state.get("formula_retry")) or any(
            entry.get("corrective") for entry in state.get("operations", [])) or any(
            count > 1 for operation, count in state.get("attempts", {}).items()
            if operation not in repeatable) or any(entry.get("corrective")
                                                  for entry in state.get("captures", []))

    def measured(self, name, function, *args, **kwargs):
        begin, started = time.monotonic(), timestamp()
        status = "blocked"
        try:
            result = function(*args, **kwargs)
            status = "passed"
            return result
        finally:
            self.state.setdefault("measurements", []).append({
                "phase": name, "pass": "qa" if self.qa_mode else "generation",
                "started_at": started, "ended_at": timestamp(), "status": status,
                "worker_active_seconds": time.monotonic() - begin})
            self.checkpoint()

    def capture_images(self, state, html_path, doc, target, out_dir, source_hash, inventory):
        if not self.modern:
            return capture_target(self.engine(), html_path, doc["format"], target, out_dir, source_hash, inventory)
        operation = f"render-{target['id']}"
        failures = [entry for entry in state.get("captures", [])
                    if entry.get("operation") == operation and entry.get("status") != "passed"]
        company.require(len(failures) < 2, f"{operation}: corrective retry already consumed")
        if failures:
            company.require(company.text(state.get("retry_reasons", {}).get(operation)),
                            f"{operation}: supply a concrete retry reason")
        entry = {"operation": operation, "status": "running", "started_at": timestamp(),
                 "pass": "qa" if self.qa_mode else "generation",
                 "artifact_sha256": company.file_hash(self.root / "documents" / doc["filename"]),
                 "corrective": bool(failures), "active_seconds": 0}
        state.setdefault("captures", []).append(entry)
        self.checkpoint()
        begin = time.monotonic()
        try:
            images, log = capture_target(self.engine(), html_path, doc["format"], target,
                                         out_dir, source_hash, inventory)
            log["artifact_sha256"] = entry["artifact_sha256"]
            company.validate_capture_result(log, self.capability)
            company.require(images and all(Path(image["path"]).is_file()
                            and Path(image["path"]).read_bytes().startswith(b"\x89PNG\r\n\x1a\n")
                            and company.file_hash(image["path"]) == image["sha256"] for image in images),
                            "Missing/unreadable/unbound capture images")
            if doc["format"] in ("docx", "pptx"):
                count_key = "pages" if doc["format"] == "docx" else "slides"
                unit_key = "page" if doc["format"] == "docx" else "slide"
                company.require(log.get("rendered_count") == target[count_key]
                                and len(images) == target[count_key]
                                and {image.get(unit_key) for image in images}
                                == set(range(1, target[count_key] + 1))
                                and (doc["format"] != "docx" or not any(i.get("overflow") for i in images)),
                                "Incomplete/overflowing page or slide capture")
            entry.update(log, status="passed", image_count=len(images))
            return images, log
        except (company.Invalid, OSError, playwright_capture.CaptureError) as exc:
            entry.update(status="blocked", blocker=str(exc))
            raise
        finally:
            entry.update(ended_at=timestamp(), elapsed_seconds=time.monotonic() - begin)
            self.checkpoint()

    @recorded_phase("document")
    def run_job(self, job):
        company.require(isinstance(job, dict) and company.text(job.get("document_id"))
                        and company.text(job.get("layout_file")), "Invalid runner job identity/layout")
        identifier = job.get("document_id")
        company.require(identifier in self.docs, "Unplanned runner job")
        doc = self.docs[identifier]
        targets = job.get("render_targets", [])
        company.require(isinstance(targets, list) and all(isinstance(t, dict) for t in targets),
                        "Invalid render targets")
        company.validate_file_ids([t.get("id") for t in targets], "render ID")
        for target in targets:
            company.portable_component(f"{identifier}-render-{target['id']}.json")
        draft_path = self.build / "drafts" / f"{identifier}.json"
        draft = company.load(draft_path)
        company.require(company.file_hash(draft_path) == self.base_selection[identifier]["draft_sha256"],
                        f"{identifier}: draft changed after execution freeze")
        company.check_draft(self.canon, self.plan, draft)
        if self.modern:
            company.require(("export_bindings" in job) != ("occurrences" in job),
                            f"{identifier}: supply exporter bindings OR explicit occurrences")
            if "export_bindings" in job:
                company.validate_export_bindings(draft, job["export_bindings"])
                command_indices = [b["command_index"] for b in job["export_bindings"] if "command_index" in b]
                if command_indices:
                    company.require("batch_file" in job, "Command bindings require batch authoring")
                    commands = company.load(batch_file(self.root, job.get("correction_file", job["batch_file"])))
                    company.require(all(i < len(commands) and commands[i]["command"] in ("add", "set")
                                        for i in command_indices),
                                    "Exporter indices must identify actual add/set commands")
            else:
                company.require(isinstance(job.get("occurrences"), list) and job["occurrences"],
                                f"{identifier}: exporter bindings or explicit occurrences required")
        layout = next(e for e in self.execution["documents"] if e["document_id"] == identifier)
        company.require(company.file_hash(local_input(self.root, job["layout_file"])) == layout["layout_sha256"],
                        f"{identifier}: changed frozen layout")
        identity = job_identity(self.root, job, self.canon, self.plan, self.modern)
        path = self.root / "documents" / doc["filename"]
        state = self.state["documents"].setdefault(identifier, {
            "input_sha256": identity, "stages": [], "operations": [], "attempts": {}})
        state["retry_reasons"] = job.get("retry_reasons", {})
        company.require(isinstance(state["retry_reasons"], dict)
                        and all(company.text(k) and company.text(v) for k, v in state["retry_reasons"].items()),
                        "Retry reasons must describe concrete corrective actions")
        company.require(state["input_sha256"] == identity,
                        f"{identifier}: changed inputs; preserve accepted output and use an explicit revision")
        if path.exists():
            company.require(state.get("artifact_sha256") == company.file_hash(path),
                            f"{identifier}: unowned/changed output; refusing overwrite")
        else:
            company.require(not state["stages"], f"{identifier}: checkpoint artifact missing")
        if "complete" in state["stages"]:
            directory = "qa-machine" if self.qa_mode else "machine"
            if self.qa_mode and not (self.build / directory / f"{identifier}.json").exists():
                directory = "machine"
            machine = company.validate_machine(self.canon, self.plan, doc, draft_path, path,
                                                self.root, self.capability, directory)
            if self.modern:
                self.persist_bindings(job, draft, machine, state)
            for entry in machine["checks"]:
                company.bound_evidence(self.root, entry, "resume")
            for render in machine["renders"]:
                for image_entry in render["images"]:
                    company.bound_evidence(self.root, image_entry, "resume")
                company.bound_evidence(self.root, render["result"], "resume")
            required = self.selection_for(identifier)
            if not required["visual_required"] or machine["renders"]:
                if self.qa_mode:
                    write_json(self.build / "qa-machine" / f"{identifier}.json", machine)
                return machine
            # Expand review evidence only; never reauthor accepted bytes.
            state["stages"].remove("complete")
            self.checkpoint()
        self.boundary()
        selected = self.selection_for(identifier)
        if not self.qa_mode and identifier not in self.execution["pilot_ids"]:
            for pilot in self.execution["pilot_ids"]:
                if self.docs[pilot]["type"] == doc["type"] and self.docs[pilot]["format"] == doc["format"]:
                    pilot_layout = next(e for e in self.execution["documents"] if e["document_id"] == pilot)
                    if pilot_layout["layout_sha256"] != layout["layout_sha256"] or (
                            pilot_layout["layout_revision"] != layout["layout_revision"]):
                        continue
                    review_path = (company.generation_review_path(self.root, pilot) if self.modern
                                   else self.build / "reviews" / f"{pilot}.json")
                    company.require(review_path.is_file(), f"{identifier}: pilot {pilot} needs actual review")
                    pilot_review = company.load(review_path)
                    pilot_machine = company.validate_machine(self.canon, self.plan, self.docs[pilot],
                        self.build / "drafts" / f"{pilot}.json",
                        self.root / "documents" / self.docs[pilot]["filename"], self.root, self.capability)
                    company.require(pilot_review.get("artifact_sha256") == pilot_machine["artifact_sha256"]
                                    and pilot_review.get("draft_sha256") == pilot_machine["draft_sha256"]
                                    and pilot_review.get("machine_sha256") == company.file_hash(
                                        self.build / "machine" / f"{pilot}.json"),
                                    "Stale pilot review")
                    company.validate_review(pilot_review, self.docs[pilot],
                                            {"content_required": True, "visual_required": True}, pilot_machine)
        path.parent.mkdir(parents=True, exist_ok=True)
        if "authored" not in state["stages"]:
            company.require(not self.qa_mode, "QA requires completed generation; it never authors missing files")
            if "batch_file" in job:
                company.require(not {"template", "data_file"} & job.keys(), "Use batch OR template/merge")
                batch = batch_file(self.root, job["batch_file"])
                if "created" not in state["stages"]:
                    self.operation(state, ["create", str(path)], "create")
                    state["stages"].append("created")
                    state["artifact_sha256"] = company.file_hash(path)
                    self.checkpoint()
                correction = (["batch", str(path), "--input", str(batch_file(self.root, job["correction_file"])),
                               "--stop-on-error"] if job.get("correction_file") else None)
                if state["attempts"].get("author"):
                    company.require(company.text(job.get("correction_reason")), "Author retry needs a concrete reason")
                    state["correction_reason"] = job["correction_reason"]
                    state["correction_sha256"] = company.file_hash(batch_file(self.root, job["correction_file"]))
                self.operation(state, ["batch", str(path), "--input", str(batch), "--stop-on-error"],
                               "author", correction)
            else:
                template = local_input(self.root, job.get("template"))
                data = local_input(self.root, job.get("data_file"))
                self.operation(state, ["merge", str(template), str(path), "--data", str(data)], "merge")
            state["stages"].append("authored")
            state["author_tool"] = self.cli.tool
            state["artifact_sha256"] = company.file_hash(path)
            self.checkpoint()
        self.boundary()
        no_page_fields(path)
        expected = self.formulas(state, path, job, draft) if doc["format"] == "xlsx" else None
        digest = company.file_hash(path)
        checks = []
        company.audit_package(path, doc["format"])
        company.verify_document_title(path, doc)
        document_log = {"status": "passed", "artifact_sha256": digest,
                        "evidence": "Checked package/self-label scans and planned visible/core business title"}
        document_path = self.build / "tool-results" / f"{identifier}-document.json"
        write_json(document_path, document_log)
        checks.append({"operation": "document", "path": document_path.relative_to(self.root).as_posix(),
                       "sha256": company.file_hash(document_path)})
        for operation, arguments in (("validate", ["validate", str(path)]),
                                     ("issues", ["view", str(path), "issues"]),
                                     ("extract", ["view", str(path), "text"])):
            log = self.inspection(state, arguments, operation, digest)
            log_path = self.build / "tool-results" / f"{identifier}-{operation}.json"
            write_json(log_path, log)
            checks.append({"operation": operation, "path": log_path.relative_to(self.root).as_posix(),
                           "sha256": company.file_hash(log_path)})
        if doc["format"] == "xlsx":
            self.boundary()
            log = self.inspection(state, ["view", str(path), "issues", "--type", "chart_cache_stale"],
                                  "chart_cache_stale", digest)
            log.update(status="passed", artifact_sha256=digest, evidence="Complete chart-cache issue scan")
            log_path = self.build / "tool-results" / f"{identifier}-chart-cache.json"
            write_json(log_path, log)
            checks.append({"operation": "chart_cache_stale", "path": log_path.relative_to(self.root).as_posix(),
                           "sha256": company.file_hash(log_path)})
            expected.update(status="passed", artifact_sha256=digest)
            expected_path = self.build / "tool-results" / f"{identifier}-formulas.json"
            write_json(expected_path, expected)
            checks.append({"operation": "formulas", "path": expected_path.relative_to(self.root).as_posix(),
                           "sha256": company.file_hash(expected_path)})
        renders = []
        selected = self.selection_for(identifier)
        if self.modern:
            verify_regions(path, targets)
        if selected["visual_required"]:
            chart_inventory = verify_regions(path, targets)
            if chart_inventory:
                # The workbook actually contains supported chart objects requiring real
                # per-chart capture; a basic cell-only region preflight probe never
                # certifies that. Require the pinned capability's own bound column/line
                # chart capture evidence (validated above by validate_preflight), not just
                # the "basic"/"html_screenshots" labels every xlsx probe always carries.
                xlsx_probe = next(p for p in self.capability["probes"] if p["format"] == "xlsx")
                company.require("chart" in xlsx_probe["features"] and xlsx_probe.get("chart_capture"),
                                f"{identifier}: this workbook has supported chart objects, but the pinned "
                                "preflight never certified chart capture capability (basic region capture "
                                "alone does not prove chart support); rerun preflight with a chart fixture")
            html_path = self.build / "renders" / identifier / "export.html"
            html_path.parent.mkdir(parents=True, exist_ok=True)
            html_log = self.inspection(state, html_export_arguments(path, html_path), "html_export", digest)
            company.require(html_path.is_file() and company.text(html_path.read_text(encoding="utf-8")),
                            "OfficeCLI HTML export missing/empty")
            html_log_path = self.build / "tool-results" / f"{identifier}-html_export.json"
            write_json(html_log_path, html_log)
            checks.append({"operation": "html_export", "path": html_log_path.relative_to(self.root).as_posix(),
                           "sha256": company.file_hash(html_log_path)})
            source_html_sha256 = company.file_hash(html_path)
            for target in targets:
                self.boundary()
                out_dir = self.build / "renders" / identifier / target["id"]
                operation = f"render-{target['id']}"
                images, capture_log = self.capture_images(state, html_path, doc, target,
                                                          out_dir, source_html_sha256, chart_inventory)
                capture_log["artifact_sha256"] = digest
                capture_log["image_count"] = len(images)
                if not self.modern:
                    state.setdefault("captures", []).append(capture_log)
                company.validate_capture_result(capture_log, self.capability)
                capture_log_path = self.build / "tool-results" / f"{identifier}-{operation}.json"
                write_json(capture_log_path, capture_log)
                image_refs = []
                for image_entry in images:
                    image_path = Path(image_entry["path"])
                    company.require(image_path.is_file()
                                    and image_path.read_bytes().startswith(b"\x89PNG\r\n\x1a\n"),
                                    "Missing/unreadable Playwright screenshot")
                    ref = {"path": image_path.relative_to(self.root).as_posix(), "sha256": image_entry["sha256"]}
                    if "page" in image_entry:
                        ref["page"] = image_entry["page"]
                        ref["overflow"] = image_entry["overflow"]
                    if "slide" in image_entry:
                        ref["slide"] = image_entry["slide"]
                    image_refs.append(ref)
                renders.append({"id": target["id"], "images": image_refs,
                    "renderer": "OfficeCLI-HTML+Playwright-Chromium", "artifact_sha256": digest,
                    "result": {"path": capture_log_path.relative_to(self.root).as_posix(),
                              "sha256": company.file_hash(capture_log_path)}})
        depth = company.load(self.build / "depth-targets.json")
        company.require(depth.get("plan_sha256") == company.fingerprint(self.plan), "Stale depth targets")
        target = next((d for d in depth["documents"] if d["document_id"] == identifier), None)
        company.require(target is not None and target["content_depth"] == doc["content_depth"],
                        "Missing/mismatched depth target")
        delivered = company.artifact_containers(path, doc["format"])
        measure = target["measure"]
        measured = company.depth_measure(delivered, measure)
        company.require(target["target_min"] <= measured <= target["target_max"], "Delivered depth target failed")
        depth_log = {"status": "passed", "artifact_sha256": digest, "measure": measure, "measured": measured,
                     "target_sha256": company.file_hash(self.build / "depth-targets.json"),
                     "evidence": "Measured delivered content against frozen targets; not substantive certification"}
        depth_path = self.build / "tool-results" / f"{identifier}-depth.json"
        write_json(depth_path, depth_log)
        checks.append({"operation": "depth", "path": depth_path.relative_to(self.root).as_posix(),
                       "sha256": company.file_hash(depth_path)})
        machine = dict(company.receipt_header(self.canon, self.plan), document_id=identifier,
            engine="OfficeCLI", tool=self.cli.tool, author_tool=state["author_tool"],
            capability=self.capability_ref, status="passed", artifact_sha256=digest,
            draft_sha256=company.file_hash(draft_path), checks=checks, renders=renders,
            render_targets=[r["id"] for r in renders],
            target_definitions=targets if renders else [], corrected_failure=self.corrected_failure(state))
        if self.modern:
            machine.update(machine_version=2, generation_sha256=company.generation_identity(self.execution),
                           declared_targets=targets, job_sha256=state["input_sha256"],
                           checker=company.checker_identity())
            self.persist_bindings(job, draft, machine, state)
            machine["export_bindings_sha256"] = company.file_hash(
                self.build / "export-bindings" / f"{identifier}.json")
        write_json(self.build / ("qa-machine" if self.qa_mode else "machine") / f"{identifier}.json", machine)
        state["stages"].append("complete")
        self.checkpoint()
        return machine

    def persist_bindings(self, job, draft, machine, state):
        identifier = job["document_id"]
        binding = dict(company.receipt_header(self.canon, self.plan),
                       document_id=identifier, artifact_sha256=machine["artifact_sha256"],
                       draft_sha256=machine["draft_sha256"], job_sha256=state["input_sha256"])
        if "export_bindings" in job:
            binding["bindings"] = job["export_bindings"]
            if "batch_file" in job:
                batch_source = job.get("correction_file") if state.get("correction_sha256") else job["batch_file"]
                binding["commands"] = company.load(batch_file(self.root, batch_source))
                logs = [entry for entry in state["operations"]
                        if entry.get("operation") == "author" and entry.get("status") == "passed"]
                company.require(logs, "Missing author result for exporter binding")
                binding["author_result"] = checked(logs[-1])
        else:
            binding["occurrences"] = job["occurrences"]
        destination = self.build / "export-bindings" / f"{identifier}.json"
        if destination.exists():
            company.require(company.load(destination) == binding,
                            f"{identifier}: exporter association changed; use an affected revision")
        else:
            write_json(destination, binding)

    @recorded_phase("assemble")
    def assemble(self, jobs):
        if self.modern and not self.qa_mode:
            return self.assemble_generation(jobs)
        company.require({j["document_id"] for j in jobs} == self.docs.keys()
                        and len(jobs) == len(self.docs), "Complete real job set required")
        manifest = dict(company.receipt_header(self.canon, self.plan), documents=[])
        locations = dict(company.receipt_header(self.canon, self.plan), documents=[])
        for job in jobs:
            identifier = job["document_id"]
            state = self.state["documents"].get(identifier, {})
            company.require("complete" in state.get("stages", [])
                            and state.get("input_sha256") == job_identity(
                                self.root, job, self.canon, self.plan, self.modern),
                            f"{identifier}: incomplete/changed job")
            path = self.root / "documents" / self.docs[identifier]["filename"]
            company.require(company.file_hash(path) == state["artifact_sha256"], "Changed final artifact")
            manifest["documents"].append({"document_id": identifier,
                "path": path.relative_to(self.root).as_posix(), "sha256": state["artifact_sha256"]})
            occurrences = job.get("occurrences")
            if self.modern:
                binding = company.load(self.build / "export-bindings" / f"{identifier}.json")
                company.require(binding["artifact_sha256"] == state["artifact_sha256"]
                                and binding["draft_sha256"] == company.file_hash(
                                    self.build / "drafts" / f"{identifier}.json")
                                and binding["job_sha256"] == state["input_sha256"],
                                f"{identifier}: stale exporter bindings")
                occurrences = (self.measured("association_resolution", company.resolve_export_bindings,
                    path, self.docs[identifier]["format"],
                    company.load(self.build / "drafts" / f"{identifier}.json"),
                    binding["bindings"], binding.get("commands"), binding.get("author_result"))
                    if "bindings" in binding else binding["occurrences"])
            locations["documents"].append({"document_id": identifier, "revision": self.docs[identifier]["revision"],
                "artifact_sha256": state["artifact_sha256"],
                "draft_sha256": company.file_hash(self.build / "drafts" / f"{identifier}.json"),
                "occurrences": occurrences})
        drafts = self.build / "drafts"
        provenance = self.measured("provenance", company.verify_provenance,
                                   self.canon, self.plan, drafts, manifest, self.root, locations)
        review_plan = self.measured("review_selection", company.make_review_plan,
                                   self.canon, self.plan, drafts, self.execution, self.root,
                                   company.manifest_files(self.canon, self.plan, manifest, self.root))
        for name, value in (("manifest", manifest), ("claim-locations", locations),
                            ("artifact-index", company.index_artifacts(self.canon, self.plan, manifest, self.root)),
                            ("provenance", provenance),
                            ("coverage", self.measured("coverage", company.coverage_report,
                                                      self.canon, self.plan, drafts, provenance)),
                            ("review-plan", review_plan)):
            destination = self.build / f"{name}.json"
            if destination.exists():
                previous = company.load(destination)
                if previous == value:
                    continue
                if name != "review-plan":
                    company.require(all(previous.get(k) == v for k, v in
                                    company.receipt_header(self.canon, self.plan).items()),
                                    f"Refusing to overwrite unrelated {name} receipt")
            write_json(destination, value)
        return review_plan


class CaptureLane:
    """Keep synchronous Playwright and its reusable browser on one owning thread."""

    def __init__(self):
        self.pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="capture")
        self.engine = None
        self.tool_identity = None

    def call(self, name, *args, **kwargs):
        queued = time.monotonic()
        def invoke():
            begin = time.monotonic()
            if self.engine is None:
                self.engine = playwright_capture.CaptureEngine()
                self.tool_identity = copy.deepcopy(self.engine.identity)
            value = getattr(self.engine, name)
            result = value(*args, **kwargs) if callable(value) else value
            if name.startswith("capture_") and isinstance(result, dict):
                result = dict(result, lane_timing={"waiting_seconds": begin - queued,
                                                   "active_seconds": time.monotonic() - begin})
            return result
        return self.pool.submit(invoke).result()

    @property
    def identity(self):
        return copy.deepcopy(self.tool_identity) if self.tool_identity is not None else self.call("identity")

    def __getattr__(self, name):
        return lambda *args, **kwargs: self.call(name, *args, **kwargs)

    def close(self):
        try:
            if self.engine is not None:
                self.pool.submit(self.engine.close).result()
        finally:
            self.pool.shutdown(wait=True)


class GenerationRunner(Runner):
    def __init__(self, root, cli):
        super().__init__(root, cli)
        company.require(self.modern, "Separate generation/QA requires a new execution_version 2 run")
        self.owner = self
        self.capture_lane = None
        self.capture_lock = Lock()
        self.blocked_families = set()
        self.wall_prior = self.state.get("active_wall_seconds", 0)
        self.validation_seen = {}
        self.worker_id = None
        phase_ids = {event["id"] for event in self.state.get("phases", [])}
        for path in (self.build / "document-state").glob("*.json"):
            company.require_build_path(path, self.root)
            checkpoint = company.load(path)
            company.require(checkpoint.get("generation_sha256") == self.state["generation_sha256"],
                            "Stale document checkpoint; create an affected revision")
            records = checkpoint.get("documents", {})
            company.require(isinstance(records, dict) and len(records) == 1,
                            "Invalid document checkpoint identity")
            identifier = next(iter(records))
            company.require(identifier in self.docs and company.same_filename(identifier, path.stem),
                            "Invalid document checkpoint identity")
            self.state["documents"][identifier] = records[identifier]
            for event in checkpoint.get("phases", []):
                if event.get("status") == "running":
                    event["status"] = "interrupted"
                if event["id"] not in phase_ids:
                    self.state.setdefault("phases", []).append(event)
                    phase_ids.add(event["id"])
        self.failure_families = {self.families[i] for i, s in self.state["documents"].items()
                                 if self.corrected_failure(s)}
        self.failure_families.update(self.families[i] for i in self.execution["flagged_ids"])

    def engine(self):
        with self.owner.capture_lock:
            if self.owner.capture_lane is None:
                self.owner.capture_lane = CaptureLane()
        return self.owner.capture_lane

    def close(self):
        if self.owner is self and self.capture_lane is not None:
            self.capture_lane.close()
            self.capture_lane = None

    def checkpoint(self):
        if self.active_phase is not None:
            self.active_phase[0]["active_seconds"] = time.monotonic() - self.active_phase[1]
        self.state["last_checkpoint_at"] = timestamp()
        if self.owner is self:
            self.state["active_wall_seconds"] = self.wall_prior + time.monotonic() - self.started
            cache = company._validation_cache.get()
            if cache:
                with cache["lock"]:
                    counts = dict(cache["counts"])
                previous = self.validation_seen.get(cache["identity"], {})
                totals = self.state.setdefault("validation_counts", {})
                for key, value in counts.items():
                    totals[key] = totals.get(key, 0) + value - previous.get(key, 0)
                self.validation_seen[cache["identity"]] = counts
        payload = self.state
        if self.worker_id is not None and self.qa_mode:
            payload = dict(self.state, documents=self.document_records())
        write_json(self.state_path, payload)

    def boundary(self):
        budget = self.execution.get("budget_seconds")
        owner = getattr(self, "owner", self)
        prior = getattr(owner, "wall_prior", self.state.get("active_wall_seconds", 0))
        company.require(budget is None or prior + time.monotonic() - self.started < budget,
                        "Runtime budget reached at a safe checkpoint; run remains partial")
        identifier = getattr(self, "worker_id", None)
        company.require(identifier is None or self.families[identifier] not in owner.blocked_families,
                        "Affected family paused at a safe checkpoint; review the recorded failure")

    def selection_for(self, identifier):
        if self.qa_mode:
            selected = dict(self.base_selection[identifier])
        else:
            required = identifier in self.execution["pilot_ids"]
            selected = {"content_required": required, "visual_required": required,
                        "reasons": ["pilot"] if required else []}
        affected = self.families[identifier] in self.owner.failure_families or any(
            self.families[i] == self.families[identifier] and self.corrected_failure(state)
            for i, state in self.state["documents"].items())
        if affected:
            selected.update(content_required=True, visual_required=True,
                            reasons=sorted(set(selected["reasons"]) | {"affected_family"}))
        return selected

    def accepted_review(self, identifier):
        review_path = company.generation_review_path(self.root, identifier)
        if not review_path.exists():
            return False
        machine = company.validate_machine(self.canon, self.plan, self.docs[identifier],
            self.build / "drafts" / f"{identifier}.json",
            self.root / "documents" / self.docs[identifier]["filename"], self.root, self.capability)
        review = company.load(review_path)
        company.require(review.get("document_id") == identifier
                        and review.get("artifact_sha256") == machine["artifact_sha256"]
                        and review.get("draft_sha256") == machine["draft_sha256"]
                        and review.get("machine_sha256") == company.file_hash(
                            self.build / "machine" / f"{identifier}.json"),
                        f"{identifier}: stale pilot/corrective review")
        company.validate_review(review, self.docs[identifier],
                                {"content_required": True, "visual_required": True}, machine)
        return True

    def worker(self, identifier):
        worker = copy.copy(self)
        worker.worker_id = identifier
        worker.queued_at = timestamp()
        worker.active_phase = None
        worker.state = {key: copy.deepcopy(value) for key, value in self.state.items()
                        if key not in ("documents", "phases")}
        worker.original_document = copy.deepcopy(self.state["documents"].get(identifier, {}))
        record = worker.original_document
        if self.qa_mode:
            record = record.get("qa_state", {k: v for k, v in record.items() if k != "qa_state"})
        worker.state["documents"] = {identifier: copy.deepcopy(record)} if record else {}
        worker.state["phases"] = []
        worker.state_path = self.build / "document-state" / f"{identifier}.json"
        return worker

    def document_records(self):
        if self.worker_id is not None and self.qa_mode:
            original = copy.deepcopy(self.original_document)
            original["qa_state"] = self.state["documents"].get(self.worker_id, {})
            return {self.worker_id: original}
        return self.state["documents"]

    def claim_snapshot(self):
        if self.state_path.exists():
            current = company.load(self.state_path)
            company.require(current.get("last_checkpoint_at") == self.state.get("last_checkpoint_at"),
                            "Runner checkpoint changed in another invocation; reopen the run")
        company.require(company.load(self.build / "execution-plan.json") == self.execution,
                        "Execution contract changed; reopen the run")

    @staticmethod
    def worker_count(jobs, requested=None):
        company.require(requested is None or company.integer(requested) and 1 <= requested <= 4,
                        "--workers must be an integer from 1 to 4")
        return max(1, min(requested or 4, os.cpu_count() or 1, len(jobs)))

    @company.validation_operation
    def run_jobs(self, jobs, workers=None):
        with run_lock(self.root):
            self.claim_snapshot()
            return self._run_jobs(jobs, workers)

    def _run_jobs(self, jobs, workers=None):
        self.blocked_families.clear()
        company.require(isinstance(jobs, list) and jobs
                        and all(isinstance(j, dict) and j.get("document_id") in self.docs for j in jobs)
                        and len({j["document_id"] for j in jobs}) == len(jobs),
                        "Jobs must contain distinct planned document IDs")
        count = self.worker_count(jobs, workers)
        diagnostics = company.family_summary(self.canon, self.plan, self.build / "drafts", self.execution, {
            i: {"corrected_failure": self.corrected_failure(s.get("qa_state", s) if self.qa_mode else s)}
            for i, s in self.state["documents"].items()})
        diagnostics["scope"] = "additional_qa" if company.execution_policy(self.execution) else "generation_pilots"
        write_json(self.build / "family-summary.json", diagnostics)
        pending = sorted(jobs, key=lambda j: (
            j["document_id"] not in self.execution["pilot_ids"],
            self.docs[j["document_id"]]["publication_date"], j["document_id"]))
        failed, awaiting, completed = {}, set(), []
        active = {}
        family_ready = {}
        unresolved = {}
        for identifier, record in self.state["documents"].items():
            stage = record.get("qa_state", record) if self.qa_mode else record
            if "complete" not in stage.get("stages", []) and (stage.get("attempts") or stage.get("captures")):
                unresolved.setdefault(self.families[identifier], set()).add(identifier)
        self.state["effective_workers"] = count
        self.checkpoint()
        try:
            with ThreadPoolExecutor(max_workers=count, thread_name_prefix="document") as pool:
                while pending or active:
                    self.boundary()
                    for job in list(pending):
                        if len(active) >= count:
                            break
                        identifier = job["document_id"]
                        family = self.families[identifier]
                        if family in self.blocked_families:
                            continue
                        if unresolved.get(family) and identifier not in unresolved[family]:
                            continue
                        if not self.qa_mode and identifier not in self.execution["pilot_ids"]:
                            pilots = [p for p in self.execution["pilot_ids"] if self.families[p] == family]
                            if family not in family_ready:
                                family_ready[family] = all(self.accepted_review(p) for p in pilots)
                            if not family_ready[family]:
                                awaiting.update(pilots)
                                continue
                        worker = self.worker(identifier)
                        active[pool.submit(copy_context().run, worker.run_job, job)] = (identifier, worker)
                        pending.remove(job)
                    if not active:
                        break
                    finished, _ = wait(active, return_when=FIRST_COMPLETED)
                    for future in finished:
                        identifier, worker = active.pop(future)
                        try:
                            machine = future.result()
                            completed.append(identifier)
                            unresolved.get(self.families[identifier], set()).discard(identifier)
                            if not self.qa_mode and machine.get("corrected_failure"):
                                self.failure_families.add(self.families[identifier])
                                if not self.accepted_review_from_worker(identifier, worker):
                                    self.blocked_families.add(self.families[identifier])
                                    awaiting.add(identifier)
                        except (company.Invalid, OSError, subprocess.SubprocessError,
                                ArithmeticError, playwright_capture.CaptureError) as exc:
                            failed[identifier] = str(exc)
                            self.blocked_families.add(self.families[identifier])
                        finally:
                            self.state["documents"].update(worker.document_records())
                            self.state.setdefault("phases", []).extend(worker.state["phases"])
                            self.checkpoint()
        finally:
            # Workers persist their own attempts before side effects, including on interruption.
            for identifier, worker in active.values():
                self.state["documents"].update(worker.document_records())
                self.state.setdefault("phases", []).extend(worker.state["phases"])
            self.checkpoint()
        corrections = sorted({i for ids in unresolved.values() for i in ids})
        result = {"status": "blocked" if failed else "awaiting_corrections" if pending and corrections else
                           "awaiting_reviews" if pending or awaiting else "passed",
                  "pass": "qa" if self.qa_mode else "generation", "workers": count,
                  "completed_this_invocation": sorted(completed), "failures": failed,
                  "generation_completed_total": sum("complete" in s.get("stages", [])
                                                     for s in self.state["documents"].values()),
                  "planned_total": len(self.docs),
                  "family_summary": {key: diagnostics[key] for key in (
                      "scope", "families", "family_sizes", "pilots", "content_selected",
                      "visual_selected", "warnings")},
                  "awaiting_pilot_or_corrective_reviews": sorted(awaiting),
                  "awaiting_corrections": corrections,
                  "remaining": sorted(j["document_id"] for j in pending)}
        self.state["last_result"] = result
        self.checkpoint()
        company.require(not failed, f"Document execution blocked: {json.dumps(result, sort_keys=True)}")
        return result

    def accepted_review_from_worker(self, identifier, worker):
        self.state["documents"].update(worker.document_records())
        return self.accepted_review(identifier)

    @company.validation_operation
    def assemble_generation(self, jobs):
        self.claim_snapshot()
        company.require({j["document_id"] for j in jobs} == self.docs.keys()
                        and len(jobs) == len(self.docs), "Complete real job set required")
        manifest = dict(company.receipt_header(self.canon, self.plan), documents=[])
        for job in sorted(jobs, key=lambda j: j["document_id"]):
            identifier = job["document_id"]
            state = self.state["documents"].get(identifier, {})
            company.require("complete" in state.get("stages", [])
                            and state.get("input_sha256") == job_identity(
                                self.root, job, self.canon, self.plan, True),
                            f"{identifier}: incomplete/changed generation job")
            manifest["documents"].append({"document_id": identifier,
                "path": (Path("documents") / self.docs[identifier]["filename"]).as_posix(),
                "sha256": state["artifact_sha256"]})
        result = self.measured("generation_verification", company.audit_generation,
                               self.canon, self.plan, self.build / "drafts", manifest, self.root)
        result.update(company.receipt_header(self.canon, self.plan),
                      generation_sha256=company.generation_identity(self.execution), completed_at=timestamp())
        write_json(self.build / "manifest.json", manifest)
        write_json(self.build / "generation.json", result)
        return result

    @company.validation_operation
    def prepare_qa(self, jobs, policy=None, workers=None):
        with run_lock(self.root):
            self.claim_snapshot()
            company.require(policy is None or policy in company.QA_POLICIES, "Unsupported QA policy")
            if policy is not None and company.execution_policy(self.execution) != policy:
                self.execution.update(revision=self.execution["revision"] + 1,
                                      post_generation_qa={"enabled": True, "policy": policy})
                company.validate_execution(self.canon, self.plan, self.execution)
                write_json(self.build / "execution-plan.json", self.execution)
                self.state.update(execution_revision=self.execution["revision"],
                                  execution_sha256=company.fingerprint(self.execution))
            company.require(company.execution_policy(self.execution) is not None,
                            "QA not requested; specify --policy for an explicitly invoked QA pass")
            manifest = company.load(self.build / "manifest.json")
            company.audit_generation(self.canon, self.plan, self.build / "drafts", manifest, self.root)
            generation_machines = {i: company.load(self.build / "machine" / f"{i}.json") for i in self.docs}
            affected = {self.families[i] for i, m in generation_machines.items() if m.get("corrected_failure")}
            affected.update(self.families[i] for i in self.execution["flagged_ids"])
            for identifier in self.docs:
                if identifier in self.execution["pilot_ids"] or self.families[identifier] in affected:
                    original = self.build / "pilot-reviews" / f"{identifier}.json"
                    if not original.exists():
                        write_json(original, company.load(company.generation_review_path(self.root, identifier)))
            self.qa_mode = True
            self.record_review_failures()
            self.base_selection = {e["document_id"]: e for e in company.review_selection(
                self.canon, self.plan, self.build / "drafts", self.execution,
                generation_machines)}
            status = dict(company.receipt_header(self.canon, self.plan), qa_status="preparing",
                          qa_policy=company.execution_policy(self.execution), started_at=timestamp())
            write_json(self.build / "qa-state.json", status)
            try:
                result = self._run_jobs(jobs, workers)
                company.require(result["status"] == "passed", "QA preparation incomplete")
                qa_machines = {i: company.load(self.build / "qa-machine" / f"{i}.json") for i in self.docs}
                corrected = {i for i, m in qa_machines.items() if m.get("corrected_failure")}
                new_flags = corrected - set(self.execution.get("qa_flagged_ids", []))
                if new_flags:
                    self.execution.update(revision=self.execution["revision"] + 1,
                        qa_flagged_ids=sorted(set(self.execution.get("qa_flagged_ids", [])) | corrected),
                        revision_reason="Expanded QA families after recorded corrective failures")
                    write_json(self.build / "execution-plan.json", self.execution)
                    self.state.update(execution_revision=self.execution["revision"],
                                      execution_sha256=company.fingerprint(self.execution))
                    self.base_selection = {e["document_id"]: e for e in company.review_selection(
                        self.canon, self.plan, self.build / "drafts", self.execution, qa_machines)}
                    self._run_jobs(jobs, workers)
                review_plan = self.assemble(jobs)
                work = self.review_work(review_plan)
                write_json(self.build / "review-work.json", work)
                status.update(qa_status="awaiting_reviews",
                              review_plan_sha256=company.fingerprint(review_plan), prepared_at=timestamp())
                write_json(self.build / "qa-state.json", status)
                return {"qa_status": status["qa_status"], "review_plan": review_plan,
                        "review_work": "build/review-work.json"}
            except (company.Invalid, OSError, subprocess.SubprocessError,
                    ArithmeticError, playwright_capture.CaptureError) as exc:
                status.update(qa_status="blocked", blocker=str(exc), ended_at=timestamp())
                write_json(self.build / "qa-state.json", status)
                raise
            finally:
                self.qa_mode = False
                self.checkpoint()

    def record_review_failures(self):
        flagged = set()
        for path in (self.build / "reviews").glob("*.json"):
            review = company.load(path)
            identifier = review.get("document_id")
            if identifier not in self.docs:
                continue
            artifact = self.root / "documents" / self.docs[identifier]["filename"]
            draft = self.build / "drafts" / f"{identifier}.json"
            if (review.get("artifact_sha256") != company.file_hash(artifact)
                    or review.get("draft_sha256") != company.file_hash(draft)):
                continue
            failed = [review.get(gate, {}) for gate in (
                "semantic_review", "genre_review", "depth_review", "visual_review", "condition_review")
                if review.get(gate, {}).get("status") in ("failed", "blocked")]
            if failed:
                company.require(all(company.text(gate.get("evidence")) for gate in failed),
                                f"{identifier}: failed review requires concrete evidence")
                finding = self.build / "qa-findings" / f"{company.fingerprint(review)}.json"
                if not finding.exists():
                    write_json(finding, review)
                flagged.add(identifier)
        if flagged - set(self.execution.get("qa_flagged_ids", [])):
            self.execution.update(revision=self.execution["revision"] + 1,
                qa_flagged_ids=sorted(set(self.execution.get("qa_flagged_ids", [])) | flagged),
                revision_reason="Expanded QA families after recorded review failures")
            write_json(self.build / "execution-plan.json", self.execution)
            self.state.update(execution_revision=self.execution["revision"],
                              execution_sha256=company.fingerprint(self.execution))
            self.checkpoint()
        return flagged

    def review_work(self, review_plan):
        previous = {}
        if (self.build / "qa.json").exists():
            previous = {r["document_id"]: r for r in company.load(self.build / "qa.json").get("documents", [])}
        result = []
        for selected in review_plan["documents"]:
            identifier = selected["document_id"]
            generation_hash = company.file_hash(self.build / "machine" / f"{identifier}.json")
            path = self.build / "reviews" / f"{identifier}.json"
            old = company.load(path) if path.exists() else previous.get(identifier, {})
            bound = old.get("artifact_sha256") == selected["artifact_sha256"] and (
                old.get("draft_sha256") == selected["draft_sha256"]) and (
                old.get("generation_machine_sha256", old.get("machine_sha256")) == generation_hash)
            review = {key: selected[key] for key in ("document_id", "artifact_sha256", "draft_sha256",
                                                    "machine_sha256")}
            review["generation_machine_sha256"] = generation_hash
            for gate in ("semantic_review", "genre_review", "depth_review", "visual_review", "condition_review"):
                required = (bool(self.docs[identifier]["condition_ids"]) if gate == "condition_review" else
                            selected["visual_required"] if gate == "visual_review" else selected["content_required"])
                reusable = bound and (gate != "visual_review"
                                      or old.get("machine_sha256") == selected["machine_sha256"])
                review[gate] = (copy.deepcopy(old[gate]) if reusable and old.get(gate, {}).get("status") == "passed"
                                else {"status": "pending"}) if required else {"status": "not_selected"}
            result.append(review)
        return dict(company.receipt_header(self.canon, self.plan),
                    execution_version=2, execution_sha256=company.fingerprint(self.execution),
                    review_plan_sha256=company.fingerprint(review_plan),
                    qa_policy=company.execution_policy(self.execution), documents=result)

    @company.validation_operation
    def finalize_qa(self):
        with run_lock(self.root):
            self.claim_snapshot()
            company.require(company.execution_policy(self.execution) is not None,
                            "Additional QA not requested")
            failures = self.record_review_failures()
            if failures:
                message = "Review failures recorded; rerun qa for expanded families: " + ", ".join(sorted(failures))
                write_json(self.build / "qa-state.json", {"qa_status": "blocked", "blocker": message})
                raise company.Invalid(message)
            files = company.manifest_files(self.canon, self.plan,
                                           company.load(self.build / "manifest.json"), self.root)
            plan = company.make_review_plan(self.canon, self.plan, self.build / "drafts",
                                           self.execution, self.root, files)
            company.require(plan == company.load(self.build / "review-plan.json"), "Stale QA preparation")
            qa = self.review_work(plan)
            for review in qa["documents"]:
                identifier = review["document_id"]
                machine = company.load(self.build / "qa-machine" / f"{identifier}.json")
                selected = next(e for e in plan["documents"] if e["document_id"] == identifier)
                company.validate_review(review, self.docs[identifier], selected, machine)
            write_json(self.build / "qa.json", qa)
            try:
                self.qa_mode = True
                with company.fresh_validation_session():
                    result = self.measured("audit", company.audit, self.canon, self.plan, self.build / "drafts",
                                           company.load(self.build / "manifest.json"), self.root)
            except (company.Invalid, OSError, ValueError) as exc:
                write_json(self.build / "qa-state.json", {"qa_status": "blocked", "blocker": str(exc)})
                raise
            finally:
                self.qa_mode = False
            result.update(qa_status="complete", completed_at=timestamp())
            write_json(self.build / "qa-state.json", result)
            return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--officecli")
    parser.add_argument("--timeout", type=int, default=120)
    commands = parser.add_subparsers(dest="command", required=True)
    probe = commands.add_parser("preflight")
    probe.add_argument("--formats", nargs="+", choices=company.FORMATS, required=True)
    probe.add_argument("--out", required=True)
    for name in ("run", "assemble", "qa"):
        command = commands.add_parser(name)
        command.add_argument("--root", required=True)
        command.add_argument("--jobs", required=True)
        if name in ("run", "qa"):
            command.add_argument("--workers", type=int)
        if name == "qa":
            command.add_argument("--policy", choices=company.QA_POLICIES)
    finalize = commands.add_parser("finalize-qa")
    finalize.add_argument("--root", required=True)
    summary = commands.add_parser("summary")
    summary.add_argument("--root", required=True)
    phase = commands.add_parser("phase")
    phase.add_argument("--root", required=True)
    phase.add_argument("--action", choices=("start", "end"), required=True)
    phase.add_argument("--name")
    phase.add_argument("--id")
    phase.add_argument("--parent")
    phase.add_argument("--pass", dest="pass_name", choices=("generation", "qa"), default="generation")
    phase.add_argument("--document")
    phase.add_argument("--status", choices=("passed", "blocked", "interrupted"), default="passed")
    args = parser.parse_args(argv)
    runner = None
    try:
        if args.command == "phase":
            result = phase_event(args.root, args.action, args.name, args.id, args.parent,
                                 args.pass_name, args.document, args.status)
            print(json.dumps(result, indent=2))
            return 0
        if args.command == "summary":
            print(json.dumps(timing_summary(args.root), indent=2))
            return 0
        cli = OfficeCLI(args.officecli, args.timeout)
        if args.command == "preflight":
            result = preflight(cli, args.formats, args.out)
        else:
            execution = company.load(Path(args.root) / "build" / "execution-plan.json")
            modern = execution.get("execution_version") == company.GENERATION_EXECUTION_VERSION
            runner = (GenerationRunner if modern else Runner)(args.root, cli)
            if args.command == "finalize-qa":
                company.require(modern, "Separate QA requires a new execution_version 2 run")
                result = runner.finalize_qa()
                print(json.dumps(result, indent=2))
                return 0
            jobs = company.load(local_input(runner.root, args.jobs, native_cli=True))
            company.require(isinstance(jobs, list) and len({j["document_id"] for j in jobs}) == len(jobs),
                            "Runner jobs must have unique document IDs")
            if args.command == "run":
                if modern:
                    result = runner.run_jobs(jobs, args.workers)
                else:
                    company.require(args.workers in (None, 1), "Existing execution_version 1 runs remain serial")
                    jobs.sort(key=lambda j: (j["document_id"] not in runner.execution["pilot_ids"],
                                            runner.docs[j["document_id"]]["publication_date"], j["document_id"]))
                    result = [runner.run_job(job) for job in jobs]
            elif args.command == "qa":
                company.require(modern, "Separate QA requires a new execution_version 2 run")
                result = runner.prepare_qa(jobs, args.policy, args.workers)
            else:
                with run_lock(args.root):
                    if modern:
                        runner.claim_snapshot()
                    result = runner.assemble(jobs)
        print(json.dumps(result, indent=2))
        return 0
    except (company.Invalid, OSError, subprocess.SubprocessError, ArithmeticError,
            playwright_capture.CaptureError) as exc:
        print(f"OfficeCLI workflow blocked: {exc}", file=sys.stderr)
        return 1
    finally:
        if runner is not None:
            runner.close()


if __name__ == "__main__":
    sys.exit(main())
