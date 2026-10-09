"""Seed, plan, and audit company content; AI authors genre-specific documents."""

import argparse
import calendar
import hashlib
import json
import random
import re
import sys
import zipfile
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree


TYPES = {
    "meeting_transcript": (["docx"], ["financial", "milestone", "acquisition"], "internal"),
    "press_release": (["docx"], ["milestone", "acquisition"], "public"),
    "product_presentation": (["pptx"], ["milestone"], "internal"),
    "product_spec": (["docx", "xlsx"], ["milestone"], "internal"),
    "quarterly_report": (["docx", "pptx", "xlsx"], ["financial"], "internal"),
    "annual_report": (["docx", "pptx", "xlsx"], ["financial"], "internal"),
    "hr_policy": (["docx"], ["policy"], "internal"),
    "board_deck": (["pptx"], ["financial", "milestone", "acquisition"], "board"),
    "financial_model": (["xlsx"], ["financial"], "internal"),
}
FORMATS = ("docx", "pptx", "xlsx")
CONTENT_DEPTHS = ("light", "medium", "deep")
DEFAULT_CONTENT_DEPTH = "medium"
MAX_FILENAME_LENGTH = 140
CONTENT_SECTIONS = {
    "meeting_transcript": ["participants", "discussion", "decisions", "actions"],
    "press_release": ["dateline_and_lead", "announcement_details", "attributed_quote",
                      "company_boilerplate", "media_contact"],
    "product_presentation": ["customer_problem", "value_proposition", "capabilities",
                             "use_cases", "next_steps"],
    "product_spec": ["scope", "requirements", "acceptance_criteria"],
    "quarterly_report": ["financial_results", "operating_review", "outlook_and_risks"],
    "annual_report": ["financial_results", "business_review", "organization", "risks"],
    "hr_policy": ["scope", "policy_rules", "administration"],
    "board_deck": ["executive_summary", "performance", "risks", "decisions_requested"],
    "financial_model": ["source_actuals", "calculations", "assumptions_and_notes"],
}


class Invalid(ValueError):
    pass


def require(condition, message):
    if not condition:
        raise Invalid(message)


def text(value):
    return isinstance(value, str) and bool(value.strip())


def integer(value):
    return type(value) is int


def iso(value):
    require(isinstance(value, str), f"Expected ISO date, got {value!r}")
    try:
        parsed = date.fromisoformat(value)
    except ValueError as exc:
        raise Invalid(f"Invalid ISO date: {value!r}") from exc
    require(parsed.isoformat() == value, f"Use YYYY-MM-DD: {value!r}")
    return parsed


def load(path):
    with Path(path).open(encoding="utf-8") as stream:
        return json.load(stream)


def save(path, value):
    path = Path(path)
    require(not path.exists(), f"Refusing to overwrite {path}; use a new revision/path")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(value, stream, indent=2, ensure_ascii=True, allow_nan=False)
        stream.write("\n")


def fingerprint(value):
    return hashlib.sha256(
        json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()
    ).hexdigest()


def file_hash(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest() if hasattr(
            hashlib, "file_digest"
        ) else hashlib.sha256(stream.read()).hexdigest()


def type_registry(request):
    registry = dict(TYPES)
    additional = request.get("additional_types", {})
    require(isinstance(additional, dict), "additional_types must be an object")
    for name, definition in additional.items():
        require(text(name) and re.fullmatch(r"[a-z][a-z0-9_]*", name) is not None
                and name not in registry, f"Invalid or duplicate additional type: {name}")
        require(isinstance(definition, dict) and set(definition) == {
            "formats", "anchor_kinds", "audience", "required_sections"},
            f"{name}: define formats, anchor_kinds, audience, and required_sections")
        formats = definition["formats"]
        require(isinstance(formats, list) and formats and all(fmt in FORMATS for fmt in formats)
                and len(formats) == len(set(formats)), f"{name}: invalid formats")
        for key in ("anchor_kinds", "required_sections"):
            values = definition[key]
            require(isinstance(values, list) and values and all(text(v) for v in values)
                    and len(values) == len(set(values)), f"{name}: invalid {key}")
        require(text(definition["audience"]), f"{name}: missing audience")
        registry[name] = (formats, definition["anchor_kinds"], definition["audience"])
    return registry


def required_sections(request, kind):
    return (CONTENT_SECTIONS[kind] if kind in CONTENT_SECTIONS
            else request["additional_types"][kind]["required_sections"])


def selected_formats(request, types):
    registry = type_registry(request)
    available = {fmt for kind in types for fmt in registry[kind][0]}
    formats = request.get("formats", [fmt for fmt in FORMATS if fmt in available])
    require(isinstance(formats, list) and formats and all(text(fmt) and fmt in available for fmt in formats)
            and len(formats) == len(set(formats)), "formats must be distinct supported Office formats")
    require(all(set(registry[kind][0]) & set(formats) for kind in types),
            "Each selected content type must support a selected format")
    require(request["documents"] >= len(formats),
            "Document target must cover each selected format")
    return formats


def validate_request(request):
    require(isinstance(request, dict), "Request must be an object")
    allowed = {"company_name", "industry", "start_year", "years", "documents",
               "seed", "document_types", "formats", "additional_types", "content_depth"}
    require(not set(request) - allowed, f"Unknown request fields: {set(request) - allowed}")
    for key in ("company_name", "industry"):
        require(text(request.get(key)), f"{key} must be nonempty text")
    for key in ("start_year", "years", "documents", "seed"):
        require(integer(request.get(key)), f"{key} must be an integer")
    require(request["start_year"] >= 2, "start_year must be at least 2")
    require(request["years"] > 0, "years must be positive")
    require(request["start_year"] + request["years"] <= 9999,
            "Period must leave room for the publication tail")
    require(request["documents"] > 0, "documents must be positive")
    require(request.get("content_depth", DEFAULT_CONTENT_DEPTH) in CONTENT_DEPTHS,
            "content_depth must be light, medium, or deep")
    registry = type_registry(request)
    types = request.get("document_types", list(TYPES))
    require(isinstance(types, list) and types and all(text(t) for t in types),
            "document_types must be a nonempty list of strings")
    require(len(types) == len(set(types)), "Duplicate document types")
    require(not set(types) - registry.keys(), f"Unsupported types: {set(types) - registry.keys()}")
    selected_formats(request, types)
    return types


def period_bounds(request):
    first = date(request["start_year"], 1, 1)
    last = date(request["start_year"] + request["years"] - 1, 12, 31)
    return first, last, last + timedelta(days=90)


def seed_canon(request):
    validate_request(request)
    request = dict(request, content_depth=request.get("content_depth", DEFAULT_CONTENT_DEPTH))
    rng = random.Random(request["seed"])
    first, last, _ = period_bounds(request)
    records = []

    def add(identifier, kind, when, data, refs=(), effective=None):
        records.append({
            "id": identifier, "kind": kind, "known_on": when.isoformat(),
            "effective_on": (effective or when).isoformat(),
            "data": data, "refs": list(refs),
        })

    add("company", "company", first, {
        "name": request["company_name"], "industry": request["industry"],
        "currency": "USD", "fiscal_year_end": "12-31",
    })
    for identifier, name in (
        ("executive", "Executive"), ("finance", "Finance"),
        ("engineering", "Engineering"), ("people", "People Operations"),
    ):
        add(f"dept-{identifier}", "department", first, {"name": name})

    people = [
        ("avery", "Avery Rowan", "CEO", "executive", None),
        ("morgan", "Morgan Vale", "CFO", "finance", "avery"),
        ("riley", "Riley Arden", "VP Engineering", "engineering", "avery"),
        ("jordan", "Jordan Ellis", "Head of People", "people", "avery"),
        ("casey", "Casey Marlow", "Finance Director", "finance", "morgan"),
        ("quinn", "Quinn Mercer", "Product Director", "engineering", "riley"),
    ]
    for identifier, name, _, _, _ in people:
        add(f"person-{identifier}", "person", first, {"name": name})
    for identifier, _, title, department, manager in people:
        refs = [f"person-{identifier}", f"dept-{department}"]
        if manager:
            refs.append(f"person-{manager}")
        add(f"role-{identifier}-initial", "role", first, {
            "person_id": f"person-{identifier}", "department_id": f"dept-{department}",
            "title": title, "reports_to": f"person-{manager}" if manager else None,
        }, refs)
    if request["years"] > 1:
        change = date(first.year + 1, 7, 1)
        add("role-casey-cfo", "role", change, {
            "person_id": "person-casey", "department_id": "dept-finance",
            "title": "CFO", "reports_to": "person-avery",
        }, ["person-casey", "dept-finance", "person-avery"])
        add("role-morgan-strategy", "role", change, {
            "person_id": "person-morgan", "department_id": "dept-executive",
            "title": "Chief Strategy Officer", "reports_to": "person-avery",
        }, ["person-morgan", "dept-executive", "person-avery"])
    for identifier, name in (("beacon", "Beacon"), ("prism", "Prism")):
        add(f"product-{identifier}", "product", first, {
            "name": name, "owner_id": "person-quinn",
        }, ["person-quinn"])
        for year in range(first.year, last.year + 1):
            version = f"{year - first.year + 1}.0"
            add(f"milestone-{identifier}-{year}-planned", "milestone", date(year, 2, 1), {
                "product_id": f"product-{identifier}", "version": version,
                "status": "planned", "target_date": f"{year}-08-15",
            }, [f"product-{identifier}"], effective=date(year, 8, 15))
            add(f"milestone-{identifier}-{year}-launched", "milestone", date(year, 8, 15), {
                "product_id": f"product-{identifier}", "version": version,
                "status": "launched",
            }, [f"product-{identifier}"])
    price = rng.randrange(8, 21) * 1_000_000
    for status, when in (("announced", date(first.year, 5, 10)),
                         ("closed", date(first.year, 9, 20))):
        add(f"acquisition-sable-{status}", "acquisition", when, {
            "acquisition_id": "deal-sable", "target": "Sableworks Analytics",
            "status": status, "price_usd": price,
        })

    revenue = rng.randrange(18, 26) * 1_000_000
    for year in range(first.year, last.year + 1):
        quarters = []
        for quarter in range(1, 5):
            revenue = revenue * rng.randrange(101, 110) // 100
            cogs = revenue * rng.randrange(25, 36) // 100
            opex = revenue * rng.randrange(35, 46) // 100
            month = quarter * 3
            end = date(year, month, calendar.monthrange(year, month)[1])
            data = {
                "year": year, "quarter": quarter, "period_end": end.isoformat(),
                "revenue_usd": revenue, "cogs_usd": cogs, "opex_usd": opex,
                "operating_income_usd": revenue - cogs - opex,
                "headcount": 140 + (year - first.year) * 24 + quarter * 6,
            }
            quarters.append(data)
            add(f"financial-{year}-q{quarter}", "financial", end + timedelta(days=15), data)
        annual = {
            "year": year, "quarter": None, "period_end": f"{year}-12-31",
            **{field: sum(q[field] for q in quarters) for field in (
                "revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd")},
            "headcount": quarters[-1]["headcount"],
        }
        add(f"financial-{year}-annual", "financial", date(year + 1, 2, 15),
            annual, [f"financial-{year}-q{q}" for q in range(1, 5)])
    for term, definition in (
        ("ARR", "Annual recurring revenue; distinct from recognized revenue."),
        ("GA", "General availability of a product version."),
        ("Operating income", "Recognized revenue less cost of goods sold and operating expenses."),
    ):
        add("glossary-" + term.lower().replace(" ", "-"), "glossary", first,
            {"term": term, "definition": definition})
    add("policy-leave-v1", "policy", first, {
        "policy_id": "paid-leave", "version": 1, "owner_id": "person-jordan",
        "annual_days": 20, "scope": "All full-time employees",
    }, ["person-jordan"])
    if request["years"] > 1:
        effective = date(first.year + 1, 1, 1)
        add("policy-leave-v2", "policy", effective - timedelta(days=30), {
            "policy_id": "paid-leave", "version": 2, "owner_id": "person-jordan",
            "annual_days": 22, "scope": "All full-time employees",
        }, ["person-jordan"], effective=effective)
    canon = {"schema_version": 1, "request": request, "records": records}
    validate_canon(canon)
    return canon


def org_chart_document(canon):
    records = canon["records"]
    index = {record["id"]: record for record in records}
    referenced = {
        ref for record in records if record["kind"] not in ("person", "department")
        for ref in record["refs"] if index[ref]["kind"] in ("person", "department")
    }
    return {
        "schema_version": 1,
        "people": [r for r in records if r["kind"] == "person" and r["id"] in referenced],
        "departments": [r for r in records if r["kind"] == "department" and r["id"] in referenced],
        "roles": [r for r in records if r["kind"] == "role"
                  and r["data"]["person_id"] in referenced],
    }


def writer_index(canon):
    chart = org_chart_document(canon)
    org_ids = {record["id"] for key in ("people", "departments", "roles") for record in chart[key]}
    return {record["id"]: record for record in canon["records"]
            if record["kind"] not in ("person", "department", "role") or record["id"] in org_ids}


def initialize_output(canon, out):
    out = Path(out)
    require(not out.exists(), f"Output directory already exists: {out}")
    save(out / "canon.json", canon)
    save(out / "build" / "request.json", canon["request"])
    (out / "documents").mkdir()


def require_build_path(path, root):
    require(Path(path).resolve().is_relative_to((Path(root) / "build").resolve()),
            f"{path}: derived files must be inside the output build directory")


def validate_output_layout(canon, root):
    root = Path(root).resolve()
    require({p.name for p in root.iterdir()} == {
        "canon.json", "documents", "build"},
        "Output root must contain only canon.json, documents, and build; "
        "legacy layouts require a new run/revision, not in-place deletion")
    require((root / "documents").is_dir() and (root / "build").is_dir(),
            "Output requires documents and build directories")
    require(all((root / folder).resolve() == root / folder for folder in ("documents", "build")),
            "Output documents/build directories may not redirect through links or junctions")
    require(fingerprint(load(root / "canon.json")) == fingerprint(canon),
            "Root canon.json differs from the audited canon")


def role_snapshot(records, as_of):
    active = {}
    for record in sorted(records, key=lambda r: (r["effective_on"], r["id"])):
        if (record["kind"] == "role" and record["known_on"] <= as_of
                and record["effective_on"] <= as_of):
            active[record["data"]["person_id"]] = record
    return active


def validate_canon(canon):
    require(isinstance(canon, dict) and canon.get("schema_version") == 1,
            "Canon must use schema_version 1")
    validate_request(canon.get("request"))
    records = canon.get("records")
    require(isinstance(records, list) and records, "Canon records must be a nonempty list")
    index = {}
    for record in records:
        require(isinstance(record, dict), "Record must be an object")
        identifier = record.get("id")
        require(text(identifier), "Record ID must be nonempty text")
        require(identifier not in index, f"Duplicate record ID: {identifier}")
        require(text(record.get("kind")), f"{identifier}: missing kind")
        iso(record.get("known_on"))
        iso(record.get("effective_on"))
        require(isinstance(record.get("data"), dict), f"{identifier}: data must be an object")
        refs = record.get("refs")
        require(isinstance(refs, list) and all(text(ref) for ref in refs),
                f"{identifier}: refs must be a string list")
        require(len(refs) == len(set(refs)), f"{identifier}: duplicate refs")
        index[identifier] = record
    require("company" in index and index["company"]["kind"] == "company",
            "Missing company record with ID company")
    company = index["company"]["data"]
    require(company.get("name") == canon["request"]["company_name"],
            "Company name must match request")
    require(company.get("currency") == "USD" and company.get("fiscal_year_end") == "12-31",
            "Version 1 supports USD and calendar fiscal years")
    first, last, _ = period_bounds(canon["request"])
    require(index["company"]["known_on"] <= first.isoformat(),
            "Company must be knowable at the start of the period")
    financials = {}
    role_dates = set()
    for record in records:
        identifier, kind, data = record["id"], record["kind"], record["data"]
        for ref in record["refs"]:
            require(ref in index and ref != identifier, f"{identifier}: invalid ref {ref}")
            require(index[ref]["known_on"] <= record["known_on"],
                    f"{identifier}: dependency {ref} is not yet knowable")

        def foreign(key, expected, nullable=False):
            ref = data.get(key)
            if nullable and ref is None:
                require(key in data, f"{identifier}: missing {key}")
                return
            require(text(ref) and ref in record["refs"] and index[ref]["kind"] == expected,
                    f"{identifier}: {key} must reference a {expected} in refs")

        if kind in ("person", "department", "product"):
            require(text(data.get("name")), f"{identifier}: missing name")
        if kind == "product":
            foreign("owner_id", "person")
        elif kind == "role":
            foreign("person_id", "person")
            foreign("department_id", "department")
            foreign("reports_to", "person", nullable=True)
            require(text(data.get("title")), f"{identifier}: missing title")
            key = (data["person_id"], record["effective_on"])
            require(key not in role_dates, f"{identifier}: ambiguous role start")
            role_dates.add(key)
        elif kind == "milestone":
            foreign("product_id", "product")
            require(text(data.get("version")) and text(data.get("status")),
                    f"{identifier}: milestone requires version and status")
            if data["status"] == "launched":
                require(record["effective_on"] <= record["known_on"],
                        f"{identifier}: a future launch cannot be recorded as completed")
        elif kind == "acquisition":
            require(all(text(data.get(key)) for key in ("acquisition_id", "target", "status")),
                    f"{identifier}: acquisition requires ID, target and status")
            require(integer(data.get("price_usd")) and data["price_usd"] >= 0,
                    f"{identifier}: invalid acquisition price")
            if data["status"] == "closed":
                require(record["effective_on"] <= record["known_on"],
                        f"{identifier}: a future close cannot be recorded as completed")
        elif kind == "policy":
            foreign("owner_id", "person")
            require(text(data.get("policy_id")) and "version" in data,
                    f"{identifier}: policy requires ID and version")
            if "name" in data:
                require(text(data["name"]), f"{identifier}: policy name must be nonempty text")
        elif kind == "glossary":
            require(text(data.get("term")) and text(data.get("definition")),
                    f"{identifier}: glossary requires term and definition")
        elif kind == "financial":
            year, quarter = data.get("year"), data.get("quarter")
            require("quarter" in data and integer(year) and first.year <= year <= last.year,
                    f"{identifier}: invalid financial year/quarter")
            require(quarter is None or (integer(quarter) and 1 <= quarter <= 4),
                    f"{identifier}: invalid quarter")
            end = iso(data.get("period_end"))
            month = (quarter or 4) * 3
            require(end == date(year, month, calendar.monthrange(year, month)[1]),
                    f"{identifier}: wrong period end")
            require(iso(record["known_on"]) > end, f"{identifier}: premature actuals")
            for field in ("revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd",
                          "headcount"):
                require(integer(data.get(field)), f"{identifier}: {field} must be integer")
                if field != "operating_income_usd":
                    require(data[field] >= 0, f"{identifier}: negative {field}")
            require(data["revenue_usd"] - data["cogs_usd"] - data["opex_usd"]
                    == data["operating_income_usd"], f"{identifier}: income does not reconcile")
            require((year, quarter) not in financials, f"{identifier}: duplicate financial period")
            financials[year, quarter] = record
    for year in range(first.year, last.year + 1):
        require(all((year, q) in financials for q in (1, 2, 3, 4, None)),
                f"{year}: requires four quarters and annual financials")
        annual = financials[year, None]
        quarters = [financials[year, q] for q in range(1, 5)]
        require(all(q["known_on"] <= annual["known_on"] for q in quarters),
                f"{year}: annual actuals precede quarterly availability")
        for field in ("revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd"):
            require(annual["data"][field] == sum(q["data"][field] for q in quarters),
                    f"{year}: annual {field} does not sum quarters")
        require(annual["data"]["headcount"] == quarters[-1]["data"]["headcount"],
                f"{year}: annual headcount must equal Q4")
    deals = {}
    for record in sorted(records, key=lambda r: (r["known_on"], r["id"])):
        if record["kind"] != "acquisition":
            continue
        data = record["data"]
        deal = deals.setdefault(data["acquisition_id"], {})
        status = data["status"]
        require(status not in deal, f"{record['id']}: duplicate acquisition stage")
        if status == "closed":
            announced = deal.get("announced")
            require(announced is not None, f"{record['id']}: close precedes announcement")
            require(announced["data"]["target"] == data["target"],
                    f"{record['id']}: acquisition target changed")
        deal[status] = record
    role_times = {max(r["known_on"], r["effective_on"]) for r in records if r["kind"] == "role"}
    for when in role_times:
        active = role_snapshot(records, when)
        for person in active:
            seen = set()
            current = person
            while current is not None:
                require(current not in seen, f"{when}: reporting cycle involving {current}")
                require(current in active, f"{when}: manager {current} has no active role")
                seen.add(current)
                current = active[current]["data"]["reports_to"]
    return index


def candidates(canon, kind):
    first, _, tail = period_bounds(canon["request"])
    anchor_kinds = type_registry(canon["request"])[kind][1]
    available = writer_index(canon)
    return sorted([
        r for r in canon["records"]
        if r["kind"] in anchor_kinds and r["id"] in available
        and first.isoformat() <= r["known_on"] <= tail.isoformat()
        and (kind != "annual_report" or r["data"].get("quarter", "missing") is None)
        and (kind not in ("quarterly_report", "financial_model")
             or r["data"].get("quarter") in (1, 2, 3, 4))
    ], key=lambda r: (r["known_on"], r["id"]))


def topic_year(record):
    return record["data"].get("year", iso(record["effective_on"]).year)


def document_title(company_name, kind, anchor, index, follow_up=0):
    data = anchor["data"]
    if anchor["kind"] == "financial":
        period = f"Q{data['quarter']} {data['year']}" if data["quarter"] else str(data["year"])
        subject = f"{period} Operating Results"
    elif anchor["kind"] == "milestone":
        product = index[data["product_id"]]["data"]["name"]
        version = f"{product} {data['version']}"
        stage = {"planned": "Release Plan", "launched": "Launch"}.get(
            data["status"], data["status"].replace("_", " ").title())
        subject = f"{version} {stage}"
    elif anchor["kind"] == "acquisition":
        stage = {"announced": "Announcement", "closed": "Close"}.get(
            data["status"], data["status"].replace("_", " ").title())
        subject = f"{data['target']} Acquisition {stage}"
    elif anchor["kind"] == "policy":
        policy = data.get("name") or data["policy_id"].replace("-", " ").replace("_", " ").title()
        if not policy.casefold().endswith("policy"):
            policy += " Policy"
        subject = f"{policy} - Version {data['version']}"
    else:
        require(text(data.get("title")), f"{anchor['id']}: naming requires a canonical title")
        subject = data["title"]

    if kind == "press_release" and anchor["kind"] == "milestone":
        action = {"planned": "Announces Plans for", "launched": "Launches"}.get(data["status"])
        title = f"{company_name} {action} {version}" if action else f"{company_name} - {subject} Update"
    elif kind == "press_release" and anchor["kind"] == "acquisition":
        action = {"announced": "Announces Acquisition of", "closed": "Completes Acquisition of"}.get(
            data["status"])
        title = f"{company_name} {action} {data['target']}" if action else f"{company_name} - {subject} Update"
    else:
        if kind == "annual_report":
            subject = f"{data['year']} Annual Business and Financial Review"
        elif kind == "financial_model":
            subject = f"{period} Financial Performance Workbook"
        elif kind == "product_spec":
            subject = f"{version} Product Specification"
        elif kind == "product_presentation":
            subject = f"{version} - Product Overview"
        elif kind == "meeting_transcript":
            subject += " - Leadership Meeting Transcript"
        elif kind == "board_deck":
            subject += " - Board Briefing"
        elif kind not in ("quarterly_report", "hr_policy"):
            subject += f" - {kind.replace('_', ' ').title()}"
        title = f"{company_name} - {subject}"
    if follow_up:
        label = "Supplement" if kind in ("quarterly_report", "annual_report") else "Follow-up"
        title += f" - {label} {follow_up}"
    return title


def document_filename(title, publication_date, fmt, revision=1, edition=1):
    cleaned = "".join(c if c.isalnum() or c in " -_." else "-" for c in title)
    cleaned = " ".join(cleaned.split()).strip(" .-")
    require(text(cleaned), "Title has no usable filename characters")
    prefix = f"{publication_date} - "
    suffix = f" - v{revision}" if revision > 1 else ""
    if edition > 1:
        suffix += f" - Edition {edition}"
    extension = f".{fmt}"
    limit = MAX_FILENAME_LENGTH - len(prefix) - len(suffix) - len(extension)
    require(limit > 3, "Filename suffix exceeds the length budget")
    if len(cleaned) > limit:
        left = (limit - 3) // 2
        cleaned = cleaned[:left].rstrip(" .-") + " - " + cleaned[-(limit - left - 3):].lstrip(" .-")
    return prefix + cleaned.rstrip(" .-") + suffix + extension


def validate_document_name(doc):
    title, filename = doc.get("title"), doc.get("filename")
    require(text(title), f"{doc['id']}: missing title")
    generic = {"document", "report", "presentation", "spreadsheet", "untitled",
               "output", "draft", doc["type"].replace("_", " ")}
    require(title.strip().casefold() not in generic
            and re.search(r"\bdoc-\d+\b", title, re.IGNORECASE) is None
            and not any(ref.casefold() in title.casefold() for ref in doc["anchor_ids"]),
            f"{doc['id']}: title must be descriptive, not a placeholder or internal ID")
    require(text(filename), f"{doc['id']}: missing filename; replan with realistic document names")
    require(len(filename) <= MAX_FILENAME_LENGTH
            and all(c.isalnum() or c in " -_." for c in filename)
            and filename == filename.strip(" .")
            and filename.startswith(doc["publication_date"] + " - ")
            and filename.endswith("." + doc["format"]),
            f"{doc['id']}: unsafe filename or missing publication date/format")
    stem = filename[len(doc["publication_date"]) + 3:-(len(doc["format"]) + 1)]
    require(stem.strip().casefold() not in generic
            and re.search(r"\bdoc-\d+\b", stem, re.IGNORECASE) is None,
            f"{doc['id']}: filename must be descriptive, not generic")


def create_plan(canon):
    index = validate_canon(canon)
    request = canon["request"]
    types = validate_request(request)
    registry = type_registry(request)
    formats = selected_formats(request, types)
    pools = {kind: candidates(canon, kind) for kind in types}
    for kind in types:
        require(pools[kind], f"No available canonical anchors for {kind}")
    documents, format_counts, type_counts = [], dict.fromkeys(formats, 0), dict.fromkeys(types, 0)
    year_counts, anchor_counts = {}, {}
    used_filenames = set()
    _, _, tail = period_bounds(request)
    for slot in range(request["documents"]):
        fmt = min(formats, key=lambda item: format_counts[item])
        eligible = [kind for kind in types if fmt in registry[kind][0]]
        kind = min(eligible, key=lambda item: (type_counts[item], types.index(item)))
        pool = pools[kind]
        anchor = min(pool, key=lambda record: (
            year_counts.get(topic_year(record), 0),
            anchor_counts.get((kind, record["id"]), 0),
            record["known_on"], record["id"],
        ))
        publication = iso(anchor["known_on"])
        cycle = anchor_counts.get((kind, anchor["id"]), 0)
        if kind not in ("quarterly_report", "annual_report"):
            publication = min(publication + timedelta(days=cycle * 7), tail)
        identifier = f"doc-{slot + 1:04d}"
        title = document_title(request["company_name"], kind, anchor, index, cycle)
        filename = document_filename(title, publication.isoformat(), fmt)
        edition = 1
        while filename.casefold() in used_filenames:
            edition += 1
            filename = document_filename(title, publication.isoformat(), fmt, edition=edition)
        used_filenames.add(filename.casefold())
        documents.append({
            "id": identifier, "type": kind, "format": fmt,
            "title": title, "filename": filename,
            "publication_date": publication.isoformat(), "audience": registry[kind][2],
            "anchor_ids": [anchor["id"]], "revision": 1,
            "content_depth": request.get("content_depth", DEFAULT_CONTENT_DEPTH),
            "required_sections": list(required_sections(request, kind)),
        })
        format_counts[fmt] += 1
        type_counts[kind] += 1
        year = topic_year(anchor)
        year_counts[year] = year_counts.get(year, 0) + 1
        anchor_counts[kind, anchor["id"]] = cycle + 1
    documents.sort(key=lambda item: (item["publication_date"], item["id"]))
    covered = {topic_year(index[d["anchor_ids"][0]]) for d in documents}
    missing_years = set(range(request["start_year"], request["start_year"] + request["years"])) - covered
    warnings = []
    if missing_years:
        warnings.append(f"No topic coverage for fiscal years: {sorted(missing_years)}")
    missing_types = [kind for kind, count in type_counts.items() if count == 0]
    if missing_types:
        warnings.append(f"Target too small to cover selected types: {missing_types}")
    if any(type_counts[kind] > len(pools[kind]) for kind in types):
        warnings.append("Repeated anchors require distinct follow-up content, not duplicate documents.")
    return {"schema_version": 1, "canon_sha256": fingerprint(canon),
            "documents": documents, "warnings": warnings}


def validate_plan(canon, plan):
    index = validate_canon(canon)
    registry = type_registry(canon["request"])
    types = validate_request(canon["request"])
    formats = set(selected_formats(canon["request"], types))
    eligible_anchors = {kind: {record["id"] for record in candidates(canon, kind)} for kind in types}
    require(isinstance(plan, dict) and plan.get("schema_version") == 1,
            "Plan must use schema_version 1")
    require(plan.get("canon_sha256") == fingerprint(canon), "Canon changed; replan in a new revision")
    documents = plan.get("documents")
    require(isinstance(documents, list) and len(documents) == canon["request"]["documents"],
            "Plan count does not match requested target")
    first, _, tail = period_bounds(canon["request"])
    by_id, filenames = {}, set()
    for doc in documents:
        require(isinstance(doc, dict) and text(doc.get("id")), "Invalid plan item")
        identifier = doc["id"]
        require(identifier not in by_id, f"Duplicate document ID: {identifier}")
        require(all(c.isalnum() or c in "-_" for c in identifier), f"Unsafe document ID: {identifier}")
        kind = doc.get("type")
        require(kind in types, f"{identifier}: unsupported type")
        require(doc.get("format") in registry[kind][0]
                and doc["format"] in formats,
                f"{identifier}: incompatible format")
        when = iso(doc.get("publication_date"))
        require(first <= when <= tail, f"{identifier}: outside publication period")
        require(doc.get("audience") == registry[kind][2], f"{identifier}: wrong audience")
        require(doc.get("required_sections") == required_sections(canon["request"], kind),
                f"{identifier}: wrong or missing genre section requirements")
        require(doc.get("content_depth") == canon["request"].get(
            "content_depth", DEFAULT_CONTENT_DEPTH),
            f"{identifier}: wrong or missing content_depth; replan in a new revision")
        require(integer(doc.get("revision")) and doc["revision"] > 0,
                f"{identifier}: invalid revision")
        anchors = doc.get("anchor_ids")
        require(isinstance(anchors, list) and anchors and all(text(ref) for ref in anchors),
                f"{identifier}: missing anchors")
        eligible = eligible_anchors[kind]
        for anchor in anchors:
            require(anchor in eligible and index[anchor]["known_on"] <= when.isoformat(),
                    f"{identifier}: invalid or future anchor {anchor}")
        validate_document_name(doc)
        require(doc["filename"].casefold() not in filenames, f"{identifier}: duplicate filename")
        filenames.add(doc["filename"].casefold())
        by_id[identifier] = doc
    require({d["format"] for d in documents} == formats,
        "Plan must contain every selected format")
    return by_id


def get_context(canon, plan, identifier):
    documents = validate_plan(canon, plan)
    require(identifier in documents, f"Unknown document: {identifier}")
    doc = documents[identifier]
    records = [r for r in writer_index(canon).values() if r["known_on"] <= doc["publication_date"]]
    active_roles = list(role_snapshot(records, doc["publication_date"]).values())
    return {
        "schema_version": 1, "canon_sha256": fingerprint(canon),
        "plan_sha256": fingerprint(plan),
        "document": doc, "company": next(r for r in records if r["id"] == "company"),
        "records": records, "active_roles": active_roles,
        "glossary": [r for r in records if r["kind"] == "glossary"],
        "org_chart": {
            "schema_version": 1,
            "people": [r for r in records if r["kind"] == "person"],
            "departments": [r for r in records if r["kind"] == "department"],
            "roles": [r for r in records if r["kind"] == "role"],
            "active_roles": active_roles,
        },
    }


def resolve(record, path):
    require(isinstance(path, list) and path, "Claim path must be a nonempty list")
    value = record
    for key in path:
        if isinstance(value, dict) and isinstance(key, str) and key in value:
            value = value[key]
        elif isinstance(value, list) and integer(key) and 0 <= key < len(value):
            value = value[key]
        else:
            raise Invalid(f"Invalid claim path: {path}")
    return value


def check_draft(canon, plan, draft):
    documents = validate_plan(canon, plan)
    index = writer_index(canon)
    _check_draft(documents, index, fingerprint(canon), fingerprint(plan), draft)


def validate_content_voice(value, label):
    wording = json.dumps(value, ensure_ascii=False) if not isinstance(value, str) else value
    forbidden = (
        r"\b(?:synthetic|fictional|fake|simulated|AI[- ]generated)\s+"
        r"(?:company|document|content|transcript|press release|report|presentation|test fixture)\b"
        r"|\b(?:this|our|the)\s+(?:company|document|content|transcript|report|presentation)"
        r"\s+(?:is|was|has been)\s+(?:an?\s+)?(?:synthetic|fictional|fake|simulated|AI[- ]generated)\b"
        r"|\b(?:generated|created|written)\s+by\s+(?:AI|an AI|artificial intelligence)\b"
        r"|\b(?:for demonstration purposes only|not a real company)\b"
        r"|^\s*(?:synthetic|fictional|fake|simulated|AI[- ]generated)\s*$"
    )
    require(re.search(forbidden, wording, re.IGNORECASE | re.MULTILINE) is None,
            f"{label}: document must not self-identify as synthetic, fictional, fake, or AI-generated")


def _check_draft(documents, index, canon_sha256, plan_sha256, draft):
    require(isinstance(draft, dict), "Draft must be an object")
    identifier = draft.get("document_id")
    require(isinstance(identifier, str) and identifier in documents, f"Unknown document: {identifier}")
    doc = documents[identifier]
    for key, value in (("canon_sha256", canon_sha256), ("plan_sha256", plan_sha256),
                       ("title", doc["title"]),
                       ("content_depth", doc["content_depth"]),
                       ("publication_date", doc["publication_date"])):
        require(type(draft.get(key)) is type(value) and draft.get(key) == value,
                f"{doc['id']}: wrong draft {key}")
    visible = {key: r for key, r in index.items() if r["known_on"] <= doc["publication_date"]}
    claims = draft.get("claims")
    require(isinstance(claims, list) and claims, f"{doc['id']}: claims required")
    by_id, cited_records = {}, set()
    for claim in claims:
        require(isinstance(claim, dict) and text(claim.get("id")), "Invalid claim")
        identifier = claim["id"]
        require(identifier not in by_id, f"Duplicate claim ID: {identifier}")
        ref = claim.get("record_id")
        require(isinstance(ref, str) and ref in visible,
                f"{identifier}: unknown or future record {ref}, or missing from the bounded canonical org graph")
        expected = resolve(visible[ref], claim.get("path"))
        require("value" in claim and fingerprint(claim["value"]) == fingerprint(expected),
                f"{identifier}: value differs from canon")
        by_id[identifier] = claim
        cited_records.add(ref)
    require(set(doc["anchor_ids"]) <= cited_records, f"{doc['id']}: uncited anchors")
    sections = draft.get("sections")
    require(isinstance(sections, list) and sections, f"{doc['id']}: sections required")
    validate_content_voice(draft["title"], doc["id"])
    used, roles = set(), set()
    for section in sections:
        require(isinstance(section, dict) and text(section.get("heading"))
                and text(section.get("text")), f"{doc['id']}: empty section")
        require(text(section.get("role")), f"{doc['id']}: section needs a genre role")
        roles.add(section["role"])
        validate_content_voice(section["heading"], doc["id"])
        validate_content_voice(section["text"], doc["id"])
        validate_content_voice(section, doc["id"])
        refs = section.get("claim_ids")
        require(isinstance(refs, list) and all(text(ref) for ref in refs),
                f"{doc['id']}: section claim_ids must be a string list")
        require(set(refs) <= by_id.keys(), f"{doc['id']}: unknown section claim")
        used.update(refs)
    require(used == by_id.keys(), f"{doc['id']}: unused claims")
    require(set(doc["required_sections"]) <= roles,
            f"{doc['id']}: missing genre sections {set(doc['required_sections']) - roles}")


def check_drafts(canon, plan, directory):
    documents = validate_plan(canon, plan)
    index = writer_index(canon)
    canon_sha256, plan_sha256 = fingerprint(canon), fingerprint(plan)
    files = list(Path(directory).glob("*.json"))
    require({p.stem for p in files} == documents.keys(), "Draft IDs do not exactly match plan")
    for path in files:
        draft = load(path)
        require(isinstance(draft, dict) and draft.get("document_id") == path.stem,
                f"{path}: filename does not match document_id")
        _check_draft(documents, index, canon_sha256, plan_sha256, draft)
    return len(files)


def audit_package(path, fmt):
    mandatory = {
        "docx": {"[Content_Types].xml", "_rels/.rels", "word/document.xml"},
        "pptx": {"[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"},
        "xlsx": {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"},
    }[fmt]
    try:
        with zipfile.ZipFile(path) as package:
            require(mandatory <= set(package.namelist()), f"{path}: missing Office parts")
            require(package.testzip() is None, f"{path}: corrupt ZIP member")
            for name in package.namelist():
                if name.endswith((".xml", ".rels")):
                    tree = ElementTree.fromstring(package.read(name))
                    for element in tree.iter():
                        if element.tag.rsplit("}", 1)[-1] == "property":
                            marker = element.attrib.get("name", "").casefold().replace("-", "_")
                            require(marker not in {"synthetic", "fictional", "fake", "ai_generated"},
                                    f"{path}: Office properties must not contain generation self-labels")
                    blocks = [element for element in tree.iter()
                              if element.tag.rsplit("}", 1)[-1] in ("p", "si", "title", "description")]
                    for block in blocks or [tree]:
                        validate_content_voice("".join(block.itertext()), f"{path}: {name}")
            if fmt == "pptx":
                require(any(name.startswith("ppt/slides/slide") and name.endswith(".xml")
                            for name in package.namelist()), f"{path}: no slides")
            if fmt == "xlsx":
                require(any(name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                            for name in package.namelist()), f"{path}: no worksheets")
    except (zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise Invalid(f"{path}: invalid Office package: {exc}") from exc


def audit(canon, plan, drafts, manifest, root):
    documents = validate_plan(canon, plan)
    check_drafts(canon, plan, drafts)
    root = Path(root).resolve()
    validate_output_layout(canon, root)
    require_build_path(drafts, root)
    require(isinstance(manifest, dict) and manifest.get("schema_version") == 1
            and manifest.get("canon_sha256") == fingerprint(canon)
            and manifest.get("plan_sha256") == fingerprint(plan), "Manifest fingerprint/version mismatch")
    entries = manifest.get("documents")
    require(isinstance(entries, list), "Manifest documents must be a list")
    seen, paths, hashes = set(), set(), {}
    for entry in entries:
        require(isinstance(entry, dict), "Manifest entry must be an object")
        identifier = entry.get("document_id")
        require(isinstance(identifier, str) and identifier in documents and identifier not in seen,
                f"Invalid/duplicate manifest document: {identifier}")
        relative = entry.get("path")
        require(text(relative) and not Path(relative).is_absolute(), f"{identifier}: path must be relative")
        path = (root / relative).resolve()
        require(path.is_relative_to(root) and path not in paths,
                f"{identifier}: duplicate or escaping artifact path")
        require(path.is_relative_to(root / "documents"),
                f"{identifier}: final artifacts must be in documents")
        require(path.is_file(), f"{identifier}: missing artifact {path}")
        fmt = documents[identifier]["format"]
        require(path.suffix == "." + fmt, f"{identifier}: wrong extension")
        require(path.name == documents[identifier]["filename"],
                f"{identifier}: artifact filename does not match the descriptive planned filename")
        digest = file_hash(path)
        require(entry.get("sha256") == digest, f"{identifier}: artifact hash mismatch")
        audit_package(path, fmt)
        seen.add(identifier)
        paths.add(path)
        hashes[identifier] = digest
    require(seen == documents.keys(), "Manifest is incomplete")
    require({p.resolve() for p in (root / "documents").rglob("*") if p.is_file()} == paths,
            "documents must contain only the manifested final artifacts; put other files in build")
    qa = load(root / "build" / "qa.json")
    require(isinstance(qa, dict) and qa.get("schema_version") == 1
            and qa.get("canon_sha256") == fingerprint(canon)
            and qa.get("plan_sha256") == fingerprint(plan), "QA fingerprint/version mismatch")
    reviews = qa.get("documents")
    require(isinstance(reviews, list), "QA documents must be a list")
    reviewed = set()
    for review in reviews:
        require(isinstance(review, dict), "Invalid QA entry")
        identifier = review.get("document_id")
        require(isinstance(identifier, str) and identifier in documents and identifier not in reviewed,
                f"Invalid/duplicate QA document: {identifier}")
        require(review.get("artifact_sha256") == hashes[identifier], f"{identifier}: stale artifact QA")
        require(review.get("draft_sha256") == file_hash(Path(drafts) / f"{identifier}.json"),
                f"{identifier}: stale draft QA")
        gates = ["semantic_review", "genre_review", "depth_review", "visual_review"]
        if documents[identifier]["format"] == "xlsx":
            gates.append("recalculation")
        for gate in gates:
            receipt = review.get(gate)
            require(isinstance(receipt, dict) and receipt.get("status") == "passed"
                    and text(receipt.get("evidence")), f"{identifier}: missing/failed {gate}")
            if gate == "depth_review":
                require(receipt.get("content_depth") == documents[identifier]["content_depth"],
                        f"{identifier}: depth_review content_depth mismatch")
        reviewed.add(identifier)
    require(reviewed == documents.keys(), "QA receipts incomplete")
    return {"documents": len(seen), "formats": {
        fmt: sum(d["format"] == fmt for d in documents.values()) for fmt in FORMATS
    }, "validation": "structure, hashes, claims, and QA receipt presence; not independent semantic certification"}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="Create a seeded baseline canon in a new directory")
    init.add_argument("--config", required=True)
    init.add_argument("--out", required=True)
    for name in ("validate", "plan", "context", "check", "audit"):
        command = commands.add_parser(name)
        command.add_argument("--canon", required=True)
        if name in ("context", "check", "audit"):
            command.add_argument("--plan", required=True)
        if name in ("plan", "context"):
            command.add_argument("--out", required=True)
        if name == "context":
            command.add_argument("--doc", required=True)
        if name in ("check", "audit"):
            command.add_argument("--drafts", required=True)
        if name == "audit":
            command.add_argument("--manifest", required=True)
            command.add_argument("--root", required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "init":
            canon = seed_canon(load(args.config))
            out = Path(args.out)
            initialize_output(canon, out)
            result = {"canon": str(out / "canon.json"),
                      "build": str(out / "build"), "records": len(canon["records"])}
        else:
            canon = load(args.canon)
            root = Path(args.canon).resolve().parent
            validate_output_layout(canon, root)
            for key in ("out", "plan", "drafts", "manifest"):
                if hasattr(args, key):
                    require_build_path(getattr(args, key), root)
            if args.command == "validate":
                result = {"records": len(validate_canon(canon)), "canon_sha256": fingerprint(canon)}
            elif args.command == "plan":
                plan = create_plan(canon)
                validate_plan(canon, plan)
                save(args.out, plan)
                result = {"documents": len(plan["documents"]), "warnings": plan["warnings"],
                          "content_depth": canon["request"].get("content_depth", DEFAULT_CONTENT_DEPTH)}
            else:
                plan = load(args.plan)
                if args.command == "context":
                    save(args.out, get_context(canon, plan, args.doc))
                    result = {"context": args.out}
                elif args.command == "check":
                    result = {"drafts": check_drafts(canon, plan, args.drafts),
                              "validation": "claim receipts only; semantic review required"}
                else:
                    result = audit(canon, plan, args.drafts, load(args.manifest), args.root)
        print(json.dumps(result, indent=2))
        return 0
    except (Invalid, OSError, ValueError, TypeError, KeyError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
