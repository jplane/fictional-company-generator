"""Checked OfficeCLI execution, capability probes, and resumable corpus mechanics."""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from datetime import datetime, timezone
from functools import wraps
from pathlib import Path
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


def recorded_phase(name):
    def decorate(function):
        @wraps(function)
        def invoke(self, *args, **kwargs):
            self.boundary()
            event = {"phase": name, "started_at": timestamp(), "status": "running", "active_seconds": 0}
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
            {"command": "set", "path": "/Results/A1", "props": {"formula": "Sheet1!A2*2"}}]


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


def local_input(root, relative):
    company.require(company.text(relative), "Missing runner input path")
    path = Path(root) / relative
    company.require_build_path(path, root)
    company.require(path.is_file(), f"Missing runner input: {relative}")
    return path


def job_identity(root, job, canon, plan):
    source_keys = ("batch_file", "template", "data_file", "layout_file")
    sources = {key: company.file_hash(local_input(root, job[key])) for key in source_keys if key in job}
    company.require("layout_file" in sources, "Every job requires its frozen layout file")
    return company.fingerprint({"canon": company.fingerprint(canon), "plan": company.fingerprint(plan),
                                "draft": company.file_hash(Path(root) / "build" / "drafts" /
                                                          f"{job['document_id']}.json"),
                                "job": {k: v for k, v in job.items()
                                        if k not in ("correction_file", "correction_reason", "retry_reasons")},
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


def capture_target(engine, html_path, fmt, target, out_dir, source_html_sha256):
    company.require(company.text(target.get("id")), "Render target requires a stable ID")
    started = timestamp()
    begin = time.monotonic()
    if fmt == "docx":
        result = engine.capture_docx_pages(html_path, out_dir, expected_pages=target["pages"])
        images = [{"page": p["page"], "path": p["path"], "sha256": p["sha256"], "overflow": p["overflow"]}
                  for p in result["pages"]]
        unit, rendered_count, extra = "pages", result["rendered_pages"], {}
    elif fmt == "pptx":
        result = engine.capture_pptx_slides(html_path, out_dir, expected_slides=target["slides"])
        images = [{"slide": s["slide"], "path": s["path"], "sha256": s["sha256"]} for s in result["slides"]]
        unit, rendered_count, extra = "slides", result["rendered_slides"], {}
    else:
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
        images = [{"path": captured["path"], "sha256": captured["sha256"]}]
        unit, rendered_count = "region", None
        extra = {"start_cell": start_cell, "end_cell": end_cell}
    log = {"operation": "capture", "status": "passed", "started_at": started, "ended_at": timestamp(),
           "active_seconds": time.monotonic() - begin, "tool": engine.identity,
           "source_html_sha256": source_html_sha256, "unit": unit, "image_sha256": images[0]["sha256"]}
    if rendered_count is not None:
        log["rendered_count"] = rendered_count
    log.update(extra)
    return images, log


def verify_regions(path, targets):
    company.verify_render_targets(path, targets)


class Runner:
    def __init__(self, root, cli):
        self.root, self.cli = Path(root).resolve(), cli
        self.build = self.root / "build"
        self.canon = company.load(self.root / "canon.json")
        self.plan = company.load(self.build / "plan.json")
        self.docs = company.validate_plan(self.canon, self.plan)
        self.execution = company.load(self.build / "execution-plan.json")
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
            "runner_version": 1, **company.receipt_header(self.canon, self.plan), "documents": {}}
        company.require(self.state.get("runner_version") == 1 and all(self.state.get(k) == v for k, v in
                        company.receipt_header(self.canon, self.plan).items()), "Stale runner canon/plan")
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
        self.capability_ref = {"path": str(capability_path.relative_to(self.root)),
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

    def operation(self, state, arguments, operation, corrective=None):
        attempts = state.setdefault("attempts", {}).get(operation, 0)
        company.require(attempts < 2, f"{operation}: corrective retry already consumed")
        if attempts:
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
                "reason": str(exc), "corrective": bool(attempts)})
            self.checkpoint()
            raise company.Invalid(f"{operation}: timed out; checkpoint preserved") from exc
        log["corrective"] = bool(attempts)
        if attempts:
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
            checked(previous[-1])
            return previous[-1]
        self.boundary()
        log = self.operation(state, arguments, operation)
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
                    if state.get("formula_retry") or any(c > 1 for c in state["attempts"].values())}
        if self.families[identifier] in affected:
            selected.update(content_required=True, visual_required=True,
                            reasons=sorted(set(selected["reasons"]) | {"affected_family"}))
        return selected

    @recorded_phase("document")
    def run_job(self, job):
        company.require(isinstance(job, dict) and company.text(job.get("document_id"))
                        and company.text(job.get("layout_file")), "Invalid runner job identity/layout")
        identifier = job.get("document_id")
        company.require(identifier in self.docs, "Unplanned runner job")
        doc = self.docs[identifier]
        draft_path = self.build / "drafts" / f"{identifier}.json"
        draft = company.load(draft_path)
        company.require(company.file_hash(draft_path) == self.base_selection[identifier]["draft_sha256"],
                        f"{identifier}: draft changed after execution freeze")
        company.check_draft(self.canon, self.plan, draft)
        layout = next(e for e in self.execution["documents"] if e["document_id"] == identifier)
        company.require(company.file_hash(local_input(self.root, job["layout_file"])) == layout["layout_sha256"],
                        f"{identifier}: changed frozen layout")
        identity = job_identity(self.root, job, self.canon, self.plan)
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
            machine = company.validate_machine(self.canon, self.plan, doc, draft_path, path,
                                                self.root, self.capability)
            for entry in machine["checks"]:
                company.bound_evidence(self.root, entry, "resume")
            for render in machine["renders"]:
                for image_entry in render["images"]:
                    company.bound_evidence(self.root, image_entry, "resume")
                company.bound_evidence(self.root, render["result"], "resume")
            required = self.selection_for(identifier)
            if not required["visual_required"] or machine["renders"]:
                return machine
            # Expand review evidence only; never reauthor accepted bytes.
            state["stages"].remove("complete")
            self.checkpoint()
        self.boundary()
        selected = self.selection_for(identifier)
        if identifier not in self.execution["pilot_ids"]:
            for pilot in self.execution["pilot_ids"]:
                if self.docs[pilot]["type"] == doc["type"] and self.docs[pilot]["format"] == doc["format"]:
                    pilot_layout = next(e for e in self.execution["documents"] if e["document_id"] == pilot)
                    if pilot_layout["layout_sha256"] != layout["layout_sha256"] or (
                            pilot_layout["layout_revision"] != layout["layout_revision"]):
                        continue
                    review_path = self.build / "reviews" / f"{pilot}.json"
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
        checks.append({"operation": "document", "path": str(document_path.relative_to(self.root)),
                       "sha256": company.file_hash(document_path)})
        for operation, arguments in (("validate", ["validate", str(path)]),
                                     ("issues", ["view", str(path), "issues"]),
                                     ("extract", ["view", str(path), "text"])):
            log = self.inspection(state, arguments, operation, digest)
            log_path = self.build / "tool-results" / f"{identifier}-{operation}.json"
            write_json(log_path, log)
            checks.append({"operation": operation, "path": str(log_path.relative_to(self.root)),
                           "sha256": company.file_hash(log_path)})
        if doc["format"] == "xlsx":
            self.boundary()
            log = self.inspection(state, ["view", str(path), "issues", "--type", "chart_cache_stale"],
                                  "chart_cache_stale", digest)
            log.update(status="passed", artifact_sha256=digest, evidence="Complete chart-cache issue scan")
            log_path = self.build / "tool-results" / f"{identifier}-chart-cache.json"
            write_json(log_path, log)
            checks.append({"operation": "chart_cache_stale", "path": str(log_path.relative_to(self.root)),
                           "sha256": company.file_hash(log_path)})
            expected.update(status="passed", artifact_sha256=digest)
            expected_path = self.build / "tool-results" / f"{identifier}-formulas.json"
            write_json(expected_path, expected)
            checks.append({"operation": "formulas", "path": str(expected_path.relative_to(self.root)),
                           "sha256": company.file_hash(expected_path)})
        targets = job.get("render_targets", [])
        company.require(isinstance(targets, list) and len({t["id"] for t in targets}) == len(targets),
                        "Invalid/duplicate render targets")
        renders = []
        corrected = bool(state.get("formula_retry")) or any(count > 1 for count in state["attempts"].values())
        selected = self.selection_for(identifier)
        if selected["visual_required"]:
            verify_regions(path, targets)
            html_path = self.build / "renders" / identifier / "export.html"
            html_path.parent.mkdir(parents=True, exist_ok=True)
            html_log = self.inspection(state, html_export_arguments(path, html_path), "html_export", digest)
            company.require(html_path.is_file() and company.text(html_path.read_text(encoding="utf-8")),
                            "OfficeCLI HTML export missing/empty")
            html_log_path = self.build / "tool-results" / f"{identifier}-html_export.json"
            write_json(html_log_path, html_log)
            checks.append({"operation": "html_export", "path": str(html_log_path.relative_to(self.root)),
                           "sha256": company.file_hash(html_log_path)})
            source_html_sha256 = company.file_hash(html_path)
            for target in targets:
                self.boundary()
                company.require(target["id"].replace("-", "").replace("_", "").isalnum(), "Unsafe render ID")
                out_dir = self.build / "renders" / identifier / target["id"]
                operation = f"render-{target['id']}"
                images, capture_log = capture_target(self.engine(), html_path, doc["format"], target,
                                                      out_dir, source_html_sha256)
                capture_log["artifact_sha256"] = digest
                company.validate_capture_result(capture_log, self.capability)
                capture_log_path = self.build / "tool-results" / f"{identifier}-{operation}.json"
                write_json(capture_log_path, capture_log)
                image_refs = []
                for image_entry in images:
                    image_path = Path(image_entry["path"])
                    company.require(image_path.is_file()
                                    and image_path.read_bytes().startswith(b"\x89PNG\r\n\x1a\n"),
                                    "Missing/unreadable Playwright screenshot")
                    ref = {"path": str(image_path.relative_to(self.root)), "sha256": image_entry["sha256"]}
                    if "page" in image_entry:
                        ref["page"] = image_entry["page"]
                        ref["overflow"] = image_entry["overflow"]
                    if "slide" in image_entry:
                        ref["slide"] = image_entry["slide"]
                    image_refs.append(ref)
                renders.append({"id": target["id"], "images": image_refs,
                    "renderer": "OfficeCLI-HTML+Playwright-Chromium", "artifact_sha256": digest,
                    "result": {"path": str(capture_log_path.relative_to(self.root)),
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
        checks.append({"operation": "depth", "path": str(depth_path.relative_to(self.root)),
                       "sha256": company.file_hash(depth_path)})
        machine = dict(company.receipt_header(self.canon, self.plan), document_id=identifier,
            engine="OfficeCLI", tool=self.cli.tool, author_tool=state["author_tool"],
            capability=self.capability_ref, status="passed", artifact_sha256=digest,
            draft_sha256=company.file_hash(draft_path), checks=checks, renders=renders,
            render_targets=[r["id"] for r in renders],
            target_definitions=targets if renders else [], corrected_failure=corrected)
        write_json(self.build / "machine" / f"{identifier}.json", machine)
        state["stages"].append("complete")
        self.checkpoint()
        return machine

    @recorded_phase("assemble")
    def assemble(self, jobs):
        company.require({j["document_id"] for j in jobs} == self.docs.keys()
                        and len(jobs) == len(self.docs), "Complete real job set required")
        manifest = dict(company.receipt_header(self.canon, self.plan), documents=[])
        locations = dict(company.receipt_header(self.canon, self.plan), documents=[])
        for job in jobs:
            identifier = job["document_id"]
            state = self.state["documents"].get(identifier, {})
            company.require("complete" in state.get("stages", [])
                            and state.get("input_sha256") == job_identity(self.root, job, self.canon, self.plan),
                            f"{identifier}: incomplete/changed job")
            path = self.root / "documents" / self.docs[identifier]["filename"]
            company.require(company.file_hash(path) == state["artifact_sha256"], "Changed final artifact")
            manifest["documents"].append({"document_id": identifier,
                "path": str(path.relative_to(self.root)), "sha256": state["artifact_sha256"]})
            locations["documents"].append({"document_id": identifier, "revision": self.docs[identifier]["revision"],
                "artifact_sha256": state["artifact_sha256"],
                "draft_sha256": company.file_hash(self.build / "drafts" / f"{identifier}.json"),
                "occurrences": job["occurrences"]})
        drafts = self.build / "drafts"
        provenance = company.verify_provenance(self.canon, self.plan, drafts, manifest, self.root, locations)
        review_plan = company.make_review_plan(self.canon, self.plan, drafts, self.execution, self.root,
                                               company.manifest_files(self.canon, self.plan, manifest, self.root))
        for name, value in (("manifest", manifest), ("claim-locations", locations),
                            ("artifact-index", company.index_artifacts(self.canon, self.plan, manifest, self.root)),
                            ("provenance", provenance),
                            ("coverage", company.coverage_report(self.canon, self.plan, drafts, provenance)),
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


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--officecli")
    parser.add_argument("--timeout", type=int, default=120)
    commands = parser.add_subparsers(dest="command", required=True)
    probe = commands.add_parser("preflight")
    probe.add_argument("--formats", nargs="+", choices=company.FORMATS, required=True)
    probe.add_argument("--out", required=True)
    for name in ("run", "assemble"):
        command = commands.add_parser(name)
        command.add_argument("--root", required=True)
        command.add_argument("--jobs", required=True)
    args = parser.parse_args(argv)
    runner = None
    try:
        cli = OfficeCLI(args.officecli, args.timeout)
        if args.command == "preflight":
            result = preflight(cli, args.formats, args.out)
        else:
            runner = Runner(args.root, cli)
            jobs = company.load(local_input(runner.root, args.jobs))
            company.require(isinstance(jobs, list) and len({j["document_id"] for j in jobs}) == len(jobs),
                            "Runner jobs must have unique document IDs")
            if args.command == "run":
                jobs.sort(key=lambda j: (j["document_id"] not in runner.execution["pilot_ids"],
                                        runner.docs[j["document_id"]]["publication_date"], j["document_id"]))
                result = [runner.run_job(job) for job in jobs]
            else:
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
