"""Seed, plan, and audit company content; AI authors genre-specific documents."""

import argparse
import calendar
import copy
import hashlib
import json
import posixpath
import random
import re
import sys
import zipfile
from collections import Counter
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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
SCHEMA_VERSION = 2
STATEMENT_KINDS = ("fact", "opinion", "assumption", "hypothesis", "recommendation",
                   "risk", "decision")
QUALIFIERS = ("uncertain", "dissent", "unresolved")
CONDITION_KINDS = ("unknown", "undocumented", "omitted_source", "conflict", "noise", "stale")
PRIVATE_KINDS = {"condition", "access_group", "knowledge_grant", "release", "viewpoint"}
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


def version(value, label):
    require(isinstance(value, dict) and value.get("schema_version") == SCHEMA_VERSION,
            f"{label} must use schema_version {SCHEMA_VERSION}; create a new run, no migration")


def string_list(value, label, nonempty=False):
    require(isinstance(value, list) and all(text(item) for item in value)
            and len(value) == len(set(value)) and (bool(value) or not nonempty),
            f"{label} must be {'nonempty ' if nonempty else ''}distinct strings")
    return value


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
               "seed", "document_types", "formats", "additional_types", "content_depth",
               "corpus_id", "viewpoints", "corpus_conditions"}
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
    if "corpus_id" in request:
        require(text(request["corpus_id"])
                and re.fullmatch(r"[A-Za-z0-9_-]+", request["corpus_id"]) is not None,
                "corpus_id must be a stable alphanumeric namespace")
    string_list(request.get("viewpoints", []), "viewpoints")
    conditions = request.get("corpus_conditions", [])
    require(isinstance(conditions, list), "corpus_conditions must be a list")
    seen = set()
    for condition in conditions:
        require(isinstance(condition, dict) and not set(condition) - {
            "id", "kind", "count", "subject_id", "document_type", "viewpoint_id", "period"},
            "Invalid condition request fields")
        require(text(condition.get("id")) and condition["id"] not in seen,
                "Condition request IDs must be nonempty and unique")
        seen.add(condition["id"])
        require(condition.get("kind") in CONDITION_KINDS, "Unsupported corpus condition kind")
        require(integer(condition.get("count")) and condition["count"] > 0,
                "Condition count must be a positive integer instance count")
        for key in ("subject_id", "viewpoint_id", "period"):
            if key in condition:
                require(text(condition[key]), f"Condition {key} must be text")
        if "document_type" in condition:
            require(condition["document_type"] in types, "Condition document_type is not selected")
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
    add("group-internal", "access_group", first,
        {"members": [f"person-{item[0]}" for item in people]},
        [f"person-{item[0]}" for item in people])
    add("group-board", "access_group", first,
        {"members": ["person-avery", "person-morgan"]}, ["person-avery", "person-morgan"])
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
    namespace = request.get("corpus_id", "corpus-" + fingerprint(request)[:24])
    canon = {"schema_version": SCHEMA_VERSION, "corpus_id": namespace,
             "request": request, "records": records}
    validate_canon(canon)
    return canon


def org_chart_document(canon):
    records = canon["records"]
    index = {record["id"]: record for record in records}
    referenced = {
        ref for record in records
        if record["kind"] not in (PRIVATE_KINDS - {"viewpoint"}) | {"person", "department"}
        for ref in record["refs"] if index[ref]["kind"] in ("person", "department")
    }
    return {
        "schema_version": SCHEMA_VERSION,
        "people": [r for r in records if r["kind"] == "person" and r["id"] in referenced],
        "departments": [r for r in records if r["kind"] == "department" and r["id"] in referenced],
        "roles": [r for r in records if r["kind"] == "role"
                  and r["data"]["person_id"] in referenced],
    }


def writer_index(canon):
    chart = org_chart_document(canon)
    org_ids = {record["id"] for key in ("people", "departments", "roles") for record in chart[key]}
    return {record["id"]: record for record in canon["records"]
            if record["kind"] not in PRIVATE_KINDS
            and (record["kind"] not in ("person", "department", "role") or record["id"] in org_ids)}


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
    version(canon, "Canon")
    require(text(canon.get("corpus_id"))
            and re.fullmatch(r"[A-Za-z0-9_-]+", canon["corpus_id"]) is not None,
            "Canon requires a stable corpus_id")
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
        string_list(record.get("access_groups", []), f"{identifier}: access_groups")
        if "knowledge_restricted" in record:
            require(type(record["knowledge_restricted"]) is bool,
                    f"{identifier}: knowledge_restricted must be boolean")
        if "public_on" in record["data"]:
            iso(record["data"]["public_on"])
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
        for group in record.get("access_groups", []):
            require(group in record["refs"] and index[group]["kind"] == "access_group",
                    f"{identifier}: access group must be declared in refs")

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
    validate_content_model(canon, index)
    return index


def validate_content_model(canon, index):
    sources = {}
    for record in index.values():
        identifier, kind, data = record["id"], record["kind"], record["data"]

        def link(key, expected=None, optional=False):
            if optional and key not in data:
                return
            ref = data.get(key)
            require(text(ref) and ref in record["refs"]
                    and (expected is None or index[ref]["kind"] == expected),
                    f"{identifier}: invalid {key} reference")

        if kind == "access_group":
            string_list(data.get("members"), f"{identifier}: members", True)
            for member in data["members"]:
                require(member in record["refs"] and index[member]["kind"] == "person",
                        f"{identifier}: group members must reference people")
        elif kind == "viewpoint":
            require(text(data.get("name")) and text(data.get("audience")),
                    f"{identifier}: viewpoint needs name and audience")
            link("author_id", "person")
            string_list(data.get("focus_ids", []), f"{identifier}: focus_ids")
            require(set(data.get("focus_ids", [])) <= set(record["refs"]),
                    f"{identifier}: viewpoint focus must be referenced")
        elif kind in ("knowledge_grant", "release"):
            link("record_id")
            if kind == "knowledge_grant":
                link("principal_id", "person")
            else:
                string_list(data.get("groups"), f"{identifier}: release groups")
                require(all(group in record["refs"] and index[group]["kind"] == "access_group"
                            for group in data["groups"]), f"{identifier}: invalid release groups")
                require(type(data.get("public", False)) is bool, f"{identifier}: public must be boolean")
                require(data["groups"] or data.get("public"), f"{identifier}: empty release")
        elif kind == "source_version":
            require(all(text(data.get(key)) for key in (
                "source_id", "title", "audience", "authority_scope")),
                f"{identifier}: source needs identity, title, audience, and authority_scope")
            require(integer(data.get("version")) and data["version"] > 0,
                    f"{identifier}: source version must be positive")
            require(data.get("status") in ("draft", "approved", "reference", "retracted"),
                    f"{identifier}: unsupported source status")
            link("author_id", "person")
            link("viewpoint_id", "viewpoint", True)
            require(data.get("document_type") in type_registry(canon["request"]),
                    f"{identifier}: source needs a supported document_type")
            require(iso(data.get("published_on")) >= iso(record["known_on"]),
                    f"{identifier}: published_on cannot precede knowability")
            key = (data["source_id"], data["version"])
            require(key not in sources, f"{identifier}: duplicate source version")
            sources[key] = record
            if "supersedes" in data:
                link("supersedes", "source_version")
                previous = index[data["supersedes"]]
                require(previous["data"]["source_id"] == data["source_id"]
                        and previous["data"]["version"] < data["version"]
                        and previous["data"]["published_on"] <= data["published_on"],
                        f"{identifier}: supersession must reference an earlier source version")
        elif kind == "statement":
            require(data.get("kind") in STATEMENT_KINDS and text(data.get("predicate"))
                    and "value" in data, f"{identifier}: invalid structured statement")
            link("subject_id")
            link("source_version_id", "source_version")
            link("speaker_id", "person")
            require(record["known_on"] <= index[data["source_version_id"]]["data"]["published_on"],
                    f"{identifier}: statement postdates its immutable source publication")
            string_list(data.get("qualifiers", []), f"{identifier}: qualifiers")
            require(set(data.get("qualifiers", [])) <= set(QUALIFIERS),
                    f"{identifier}: invalid statement qualifier")
            string_list(data.get("supports", []), f"{identifier}: supports")
            require(set(data.get("supports", [])) <= set(record["refs"]),
                    f"{identifier}: supports must be declared in refs")
            for key in ("corrects", "retracts"):
                link(key, "statement", True)
                if key in data:
                    previous = index[data[key]]["data"]
                    require(index[data[key]]["known_on"] <= record["known_on"]
                            and data[key] != identifier, f"{identifier}: invalid backward lifecycle link")
                    require(all(previous.get(field) == data.get(field) for field in (
                        "subject_id", "predicate", "period", "units")),
                        f"{identifier}: lifecycle statements must have the same scope")
            for key in ("units", "period"):
                if key in data:
                    require(text(data[key]), f"{identifier}: {key} must be text")
        elif kind == "condition":
            allowed = {"request_id", "kind", "subject_id", "statement_ids", "source_version_ids",
                       "blocked_record_ids", "as_of_date", "document_type", "viewpoint_id", "period"}
            require(not set(data) - allowed, f"{identifier}: unknown condition fields")
            require(data.get("kind") in CONDITION_KINDS and text(data.get("request_id")),
                    f"{identifier}: invalid corpus condition")
            link("subject_id")
            link("viewpoint_id", "viewpoint", True)
            if "document_type" in data:
                require(data["document_type"] in type_registry(canon["request"]),
                        f"{identifier}: unsupported condition document_type")
            if "period" in data:
                require(text(data["period"]), f"{identifier}: period must be text")
            for key, expected in (("statement_ids", "statement"),
                                  ("source_version_ids", "source_version"),
                                  ("blocked_record_ids", None)):
                string_list(data.get(key, []), f"{identifier}: {key}")
                for ref in data.get(key, []):
                    require(ref in record["refs"]
                            and (expected is None or index[ref]["kind"] == expected),
                            f"{identifier}: invalid {key}")
            if data["kind"] == "conflict":
                require(len(data.get("statement_ids", [])) >= 2,
                        f"{identifier}: conflict requires at least two statements")
                assertions = [index[ref]["data"] for ref in data["statement_ids"]]
                keys = {(item["subject_id"], item["predicate"], item.get("period"), item.get("units"))
                        for item in assertions}
                require(len(keys) == 1 and len({fingerprint(item["value"]) for item in assertions}) > 1,
                        f"{identifier}: conflict must concern differing assertions in the same scope")
                require(data["subject_id"] == assertions[0]["subject_id"]
                        and ("period" not in data or data["period"] == assertions[0].get("period")),
                        f"{identifier}: conflict subject/period differs from its assertions")
            elif data["kind"] in ("unknown", "undocumented"):
                require(data.get("blocked_record_ids"), f"{identifier}: gap needs blocked records")
            elif data["kind"] in ("omitted_source", "noise", "stale"):
                require(data.get("source_version_ids"), f"{identifier}: condition needs source versions")
            if data["kind"] == "stale":
                as_of = iso(data.get("as_of_date"))
                for ref in data["source_version_ids"]:
                    source = index[ref]
                    require(any(newer["kind"] == "source_version"
                                and newer["data"].get("supersedes") == ref
                                and iso(newer["data"]["published_on"]) <= as_of
                                for newer in index.values()),
                            f"{identifier}: stale source needs a published superseding version")

    requested = {item["id"]: item for item in canon["request"].get("corpus_conditions", [])}
    conflicts = set()
    for record in index.values():
        if record["kind"] != "condition":
            continue
        data = record["data"]
        require(data["request_id"] in requested
                and requested[data["request_id"]]["kind"] == data["kind"],
                f"{record['id']}: condition is not explicitly requested")
        for key in ("subject_id", "document_type", "viewpoint_id", "period"):
            if key in requested[data["request_id"]]:
                require(data.get(key) == requested[data["request_id"]][key],
                        f"{record['id']}: condition selector {key} differs from request")
        if data["kind"] == "conflict":
            ids = data["statement_ids"]
            conflicts.update(frozenset((a, b)) for a in ids for b in ids if a != b)
    facts = {}
    def predecessors(identifier, fields):
        result, pending = set(), [identifier]
        while pending:
            current = pending.pop()
            for field in fields:
                ref = index[current]["data"].get(field)
                if ref:
                    require(ref != identifier, f"{identifier}: cyclic source/statement lifecycle")
                    if ref not in result:
                        result.add(ref)
                        pending.append(ref)
        return result
    histories = {record["id"]: predecessors(record["id"], ("corrects", "retracts"))
                 for record in index.values() if record["kind"] == "statement"}
    source_histories = {record["id"]: predecessors(record["id"], ("supersedes",))
                        for record in index.values() if record["kind"] == "source_version"}
    for record in index.values():
        if record["kind"] != "statement" or record["data"]["kind"] != "fact":
            continue
        data = record["data"]
        key = (data["subject_id"], data["predicate"], data.get("period"), data.get("units"))
        for previous in facts.get(key, []):
            if fingerprint(previous["data"]["value"]) == fingerprint(data["value"]):
                continue
            lifecycle = previous["id"] in histories[record["id"]] or record["id"] in histories[previous["id"]]
            old_source = index[previous["data"]["source_version_id"]]
            new_source = index[data["source_version_id"]]
            superseded = (old_source["id"] in source_histories[new_source["id"]]
                          or new_source["id"] in source_histories[old_source["id"]])
            require(lifecycle or superseded or frozenset((previous["id"], record["id"])) in conflicts,
                    f"{record['id']}: unregistered conflicting factual assertion")
        facts.setdefault(key, []).append(record)


def principal_groups(index, principal, as_of):
    return {record["id"] for record in index.values()
            if record["kind"] == "access_group" and record["known_on"] <= as_of
            and record["effective_on"] <= as_of and principal in record["data"]["members"]}


def accessible(record, index, principal, as_of):
    if record["known_on"] > as_of:
        return False
    groups = principal_groups(index, principal, as_of)
    restricted = set(record.get("access_groups", []))
    released = any(item["kind"] == "release" and item["known_on"] <= as_of
                   and item["effective_on"] <= as_of and item["data"]["record_id"] == record["id"]
                   and (item["data"].get("public") or groups & set(item["data"]["groups"]))
                   for item in index.values())
    if restricted and not restricted & groups and not released:
        return False
    public_on = record["data"].get("public_on")
    return not (principal == "public" and public_on is not None and public_on > as_of)


def project_records(index, doc, include_knowledge=True, for_claims=False):
    as_of, principal = doc["publication_date"], doc["author_id"]
    denied = set(doc.get("blocked_record_ids", []))
    if for_claims:
        denied.update(doc.get("suppressed_record_ids", []))
    chart = org_chart_document({"records": list(index.values())})
    org_ids = {item["id"] for items in (chart["people"], chart["departments"], chart["roles"])
               for item in items}
    visible = {}
    for identifier, record in index.items():
        if record["kind"] in PRIVATE_KINDS or identifier in denied:
            continue
        if record["kind"] in ("person", "department", "role") and identifier not in org_ids:
            continue
        if not accessible(record, index, principal, as_of):
            continue
        if include_knowledge and record.get("knowledge_restricted", False):
            grants = [item for item in index.values()
                      if item["kind"] == "knowledge_grant"
                      and item["data"]["record_id"] == identifier
                      and item["data"]["principal_id"] == principal
                      and item["known_on"] <= as_of and item["effective_on"] <= as_of]
            if not grants:
                continue
        visible[identifier] = record
    changed = True
    while changed:
        changed = False
        for identifier, record in list(visible.items()):
            business_refs = {ref for ref in record["refs"] if index[ref]["kind"] not in PRIVATE_KINDS}
            if not business_refs <= visible.keys():
                del visible[identifier]
                changed = True
    return visible


def candidates(canon, kind):
    first, _, tail = period_bounds(canon["request"])
    anchor_kinds = type_registry(canon["request"])[kind][1]
    available = writer_index(canon)
    omitted = {ref for record in canon["records"] if record["kind"] == "condition"
               and record["data"]["kind"] == "omitted_source"
               for ref in record["data"].get("source_version_ids", [])}
    return sorted([
        r for r in canon["records"]
        if (r["kind"] in anchor_kinds or (r["kind"] == "source_version"
            and r["data"]["document_type"] == kind)) and r["id"] in available and r["id"] not in omitted
        and first.isoformat() <= r["known_on"] <= tail.isoformat()
        and (r["kind"] == "source_version"
             or kind != "annual_report" or r["data"].get("quarter", "missing") is None)
        and (r["kind"] == "source_version" or kind not in ("quarterly_report", "financial_model")
             or r["data"].get("quarter") in (1, 2, 3, 4))
    ], key=lambda r: (r["known_on"], r["id"]))


def topic_year(record):
    return record["data"].get("year", iso(record["effective_on"]).year)


def document_title(company_name, kind, anchor, index, follow_up=0):
    data = anchor["data"]
    if anchor["kind"] == "source_version":
        title = f"{company_name} - {data['title']}"
        return title + (f" - Follow-up {follow_up}" if follow_up else "")
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


def condition_instances(canon):
    instances = [record for record in canon["records"] if record["kind"] == "condition"]
    counts = Counter(item["data"]["request_id"] for item in instances)
    for request in canon["request"].get("corpus_conditions", []):
        require(counts[request["id"]] == request["count"],
                f"Condition {request['id']}: expected {request['count']} canonical instances, "
                f"found {counts[request['id']]}; enrich the canon before planning")
    return instances


def condition_sources(record, index):
    data = record["data"]
    return set(data.get("source_version_ids", [])) | {
        index[ref]["data"]["source_version_id"] for ref in data.get("statement_ids", [])}


def document_profile(canon, anchor, viewpoint_id=None):
    index = {record["id"]: record for record in canon["records"]}
    source = anchor["data"] if anchor["kind"] == "source_version" else {}
    viewpoint_id = source.get("viewpoint_id", viewpoint_id)
    profile = index[viewpoint_id]["data"] if viewpoint_id else {}
    author = source.get("author_id", profile.get("author_id", "person-avery"))
    audience = source.get("audience", profile.get("audience"))
    if profile and source:
        require(author == profile["author_id"] and audience == profile["audience"],
                f"{anchor['id']}: source and viewpoint author/audience differ")
    return author, audience, viewpoint_id


def document_access_groups(anchor, audience):
    if anchor["kind"] == "source_version" and "access_groups" in anchor:
        return anchor["access_groups"]
    return [] if audience == "public" else [
        "group-board" if audience == "board" else "group-internal"]


def document_periods(doc, index):
    covered = set()
    for ref in doc["anchor_ids"]:
        record = index[ref]
        data = record["data"]
        if record["kind"] == "financial":
            covered.add(str(data["year"]) + (f"-Q{data['quarter']}" if data.get("quarter") else ""))
        else:
            covered.add(str(topic_year(record)))
    if doc.get("source_version_id"):
        covered.update(record["data"]["period"] for record in index.values()
                       if record["kind"] == "statement"
                       and record["data"]["source_version_id"] == doc["source_version_id"]
                       and "period" in record["data"])
    return covered


def assign_conditions(canon, documents):
    index = {record["id"]: record for record in canon["records"]}
    assignments = {doc["id"]: {"condition_ids": [], "blocked_record_ids": [],
                               "suppressed_record_ids": []} for doc in documents}
    for condition in condition_instances(canon):
        data = condition["data"]
        sources = condition_sources(condition, index)
        matches = [doc for doc in documents if (
            (sources and doc.get("source_version_id") in sources)
            or (not sources and any(
                ref == data["subject_id"] or data["subject_id"] in index[ref]["refs"]
                for ref in doc["anchor_ids"])))]
        if data["kind"] == "omitted_source":
            require(not any(doc.get("source_version_id") in sources for doc in documents),
                    f"{condition['id']}: omitted source was planned")
            matches = [doc for doc in documents if any(
                ref == data["subject_id"] or data["subject_id"] in index[ref]["refs"]
                for ref in doc["anchor_ids"])]
        matches = [doc for doc in matches
                   if ("document_type" not in data or doc["type"] == data["document_type"])
                   and ("viewpoint_id" not in data or doc["viewpoint_id"] == data["viewpoint_id"])
                   and ("period" not in data or data["period"] in document_periods(doc, index))]
        require(matches, f"{condition['id']}: no compatible carrier document in the fixed budget")
        if data["kind"] in ("conflict", "noise", "stale"):
            require(sources <= {doc.get("source_version_id") for doc in matches},
                    f"{condition['id']}: not all required sources were planned")
        for doc in matches:
            assignments[doc["id"]]["condition_ids"].append(condition["id"])
            if data["kind"] in ("unknown", "undocumented"):
                key = "blocked_record_ids" if data["kind"] == "unknown" else "suppressed_record_ids"
                assignments[doc["id"]][key].extend(data["blocked_record_ids"])
    for doc in documents:
        forbidden = set(assignments[doc["id"]]["blocked_record_ids"]) | set(
            assignments[doc["id"]]["suppressed_record_ids"])
        required_conflicts = {ref for condition in canon["records"] if condition["kind"] == "condition"
                              and condition["data"]["kind"] == "conflict"
                              and condition["id"] in assignments[doc["id"]]["condition_ids"]
                              for ref in condition["data"]["statement_ids"]
                              if index[ref]["data"]["source_version_id"] == doc["source_version_id"]}
        require(not forbidden & required_conflicts,
                f"{doc['id']}: gap and conflict requirements overlap incompatibly")
    return {identifier: {key: sorted(set(value)) for key, value in item.items()}
            for identifier, item in assignments.items()}


def create_plan(canon):
    index = validate_canon(canon)
    request = canon["request"]
    types = validate_request(request)
    registry = type_registry(request)
    formats = selected_formats(request, types)
    pools = {kind: candidates(canon, kind) for kind in types}
    for kind in types:
        require(pools[kind], f"No available canonical anchors for {kind}")
    viewpoints = request.get("viewpoints", [])
    require(all(ref in index and index[ref]["kind"] == "viewpoint" for ref in viewpoints),
            "Requested viewpoints must exist in the canon before planning")
    instances = condition_instances(canon)
    mandatory_sources = sorted({ref for record in instances
                                if record["data"]["kind"] != "omitted_source"
                                for ref in condition_sources(record, index)},
                               key=lambda ref: (index[ref]["data"]["published_on"], ref))
    mandatory = []
    for ref in mandatory_sources:
        anchor = index[ref]
        kind = anchor["data"]["document_type"]
        require(kind in pools and anchor in pools[kind],
                f"{ref}: requested condition source is incompatible, omitted, or outside the period")
        mandatory.append((kind, anchor, anchor["data"].get("viewpoint_id")))
    for viewpoint in viewpoints:
        if any(profile == viewpoint for _, _, profile in mandatory):
            continue
        matches = [(kind, anchor, viewpoint) for kind in types for anchor in pools[kind]
                   if anchor["kind"] != "source_version"
                   or anchor["data"].get("viewpoint_id") == viewpoint]
        require(matches, f"{viewpoint}: no compatible viewpoint anchor")
        mandatory.append(matches[0])
    require(len(mandatory) <= request["documents"],
            "Required sources/viewpoints exceed the fixed document budget")
    documents, format_counts, type_counts = [], dict.fromkeys(formats, 0), dict.fromkeys(types, 0)
    year_counts, anchor_counts = {}, {}
    used_filenames = set()
    _, _, tail = period_bounds(request)
    for slot in range(request["documents"]):
        if slot < len(mandatory):
            kind, anchor, viewpoint_id = mandatory[slot]
            eligible_formats = [fmt for fmt in formats if fmt in registry[kind][0]]
            fmt = min(eligible_formats, key=lambda item: format_counts[item])
        else:
            fmt = min(formats, key=lambda item: format_counts[item])
            eligible = [kind for kind in types if fmt in registry[kind][0]]
            kind = min(eligible, key=lambda item: (type_counts[item], types.index(item)))
            anchor = min(pools[kind], key=lambda record: (
                year_counts.get(topic_year(record), 0),
                anchor_counts.get((kind, record["id"]), 0),
                record["known_on"], record["id"],
            ))
            viewpoint_id = viewpoints[slot % len(viewpoints)] if viewpoints else None
        source_id = anchor["id"] if anchor["kind"] == "source_version" else None
        publication = iso(anchor["data"]["published_on"] if source_id else anchor["known_on"])
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
        author, audience, viewpoint_id = document_profile(canon, anchor, viewpoint_id)
        audience = audience or registry[kind][2]
        groups = document_access_groups(anchor, audience)
        documents.append({
            "id": identifier, "type": kind, "format": fmt,
            "title": title, "filename": filename,
            "publication_date": publication.isoformat(), "audience": audience,
            "anchor_ids": [anchor["id"]], "revision": 1,
            "author_id": author, "viewpoint_id": viewpoint_id,
            "access_groups": groups, "source_version_id": source_id,
            "required_statement_ids": sorted(record["id"] for record in index.values()
                                             if record["kind"] == "statement"
                                             and record["data"]["source_version_id"] == source_id
                                             and record["known_on"] <= publication.isoformat()),
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
    assignments = assign_conditions(canon, documents)
    for doc in documents:
        doc.update(assignments[doc["id"]])
        doc["required_statement_ids"] = [ref for ref in doc["required_statement_ids"]
                                         if ref not in doc["blocked_record_ids"] + doc["suppressed_record_ids"]]
    plan = {"schema_version": SCHEMA_VERSION, "corpus_id": canon["corpus_id"],
            "canon_sha256": fingerprint(canon), "documents": documents, "warnings": warnings}
    validate_plan(canon, plan)
    return plan


def validate_plan(canon, plan):
    index = validate_canon(canon)
    registry = type_registry(canon["request"])
    types = validate_request(canon["request"])
    formats = set(selected_formats(canon["request"], types))
    eligible_anchors = {kind: {record["id"] for record in candidates(canon, kind)} for kind in types}
    version(plan, "Plan")
    require(plan.get("corpus_id") == canon["corpus_id"], "Plan corpus identity mismatch")
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
        source_id = doc.get("source_version_id")
        expected_source = anchors[0] if index[anchors[0]]["kind"] == "source_version" else None
        require(source_id == expected_source, f"{identifier}: wrong source_version_id")
        viewpoint_id = doc.get("viewpoint_id")
        require(viewpoint_id is None or (viewpoint_id in index
                and index[viewpoint_id]["kind"] == "viewpoint"
                and index[viewpoint_id]["known_on"] <= when.isoformat()),
                f"{identifier}: unknown or future viewpoint")
        author, audience, _ = document_profile(canon, index[anchors[0]], viewpoint_id)
        require(doc.get("author_id") == author and author in index
                and index[author]["kind"] == "person", f"{identifier}: wrong author")
        require(doc.get("audience") == (audience or registry[kind][2]), f"{identifier}: wrong audience")
        groups = document_access_groups(index[anchors[0]], doc["audience"])
        require(doc.get("access_groups") == groups, f"{identifier}: wrong intended access groups")
        require(doc["audience"] != "public" or not groups,
                f"{identifier}: public documents cannot have restricted intended access")
        require(all(group in index and index[group]["kind"] == "access_group" for group in groups),
                f"{identifier}: missing access groups")
        if source_id:
            require(index[source_id]["data"]["published_on"] <= when.isoformat(),
                    f"{identifier}: unpublished source")
        visible = project_records(index, doc, for_claims=True)
        require(set(anchors) <= visible.keys() and author in visible,
                f"{identifier}: anchor/author invisible to the author; revise the plan or knowledge grants")
        if viewpoint_id:
            require(set(index[viewpoint_id]["data"].get("focus_ids", [])) <= visible.keys(),
                    f"{identifier}: viewpoint focus is invisible")
        for ref in doc.get("required_statement_ids", []):
            require(ref in visible and index[ref]["kind"] == "statement"
                    and index[ref]["data"]["source_version_id"] == source_id,
                    f"{identifier}: required statement is invisible or has wrong source")
        expected_statements = sorted(record["id"] for record in index.values()
                                     if record["kind"] == "statement"
                                     and record["data"]["source_version_id"] == source_id
                                     and record["known_on"] <= when.isoformat()
                                     and record["id"] not in doc.get("blocked_record_ids", [])
                                     + doc.get("suppressed_record_ids", []))
        require(doc.get("required_statement_ids") == expected_statements,
                f"{identifier}: wrong required_statement_ids")
        validate_document_name(doc)
        require(doc["filename"].casefold() not in filenames, f"{identifier}: duplicate filename")
        filenames.add(doc["filename"].casefold())
        by_id[identifier] = doc
    require({d["format"] for d in documents} == formats,
        "Plan must contain every selected format")
    require(set(canon["request"].get("viewpoints", [])) <= {
        doc["viewpoint_id"] for doc in documents}, "Requested viewpoint coverage is incomplete")
    assignments = assign_conditions(canon, documents)
    for doc in documents:
        for key, expected in assignments[doc["id"]].items():
            require(doc.get(key) == expected, f"{doc['id']}: wrong {key}")
    return by_id


def writer_record(record, projected):
    result = copy.deepcopy(record)
    result["refs"] = [ref for ref in record["refs"] if ref in projected]
    result.pop("access_groups", None)
    result.pop("knowledge_restricted", None)
    return result


def get_context(canon, plan, identifier):
    documents = validate_plan(canon, plan)
    require(identifier in documents, f"Unknown document: {identifier}")
    doc = documents[identifier]
    index = {record["id"]: record for record in canon["records"]}
    projected = project_records(index, doc)
    records = [writer_record(record, projected) for record in projected.values()]
    writer_doc = {key: value for key, value in doc.items()
                  if key not in ("condition_ids", "blocked_record_ids", "access_groups")}
    viewpoint = None
    if doc["viewpoint_id"]:
        profile = index[doc["viewpoint_id"]]["data"]
        viewpoint = {key: copy.deepcopy(profile[key]) for key in ("name", "author_id", "audience")}
        viewpoint["focus_ids"] = [ref for ref in profile.get("focus_ids", []) if ref in projected]
    active_roles = list(role_snapshot(records, doc["publication_date"]).values())
    return {
        "schema_version": SCHEMA_VERSION, "corpus_id": canon["corpus_id"],
        "canon_sha256": fingerprint(canon),
        "plan_sha256": fingerprint(plan),
        "document": writer_doc, "company": next(r for r in records if r["id"] == "company"),
        "viewpoint": viewpoint,
        "records": records, "active_roles": active_roles,
        "glossary": [r for r in records if r["kind"] == "glossary"],
        "org_chart": {
            "schema_version": SCHEMA_VERSION,
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


def evaluate_claim(identifier, claims, visiting=None):
    visiting = set() if visiting is None else visiting
    require(identifier in claims and identifier not in visiting,
            f"{identifier}: unknown input or cyclic derivation")
    claim = claims[identifier]
    if "derivation" not in claim:
        return claim["value"]
    require(not set(claim) & {"record_id", "path", "statement_kind"},
            f"{identifier}: derived claim cannot also be a primary record assertion")
    definition = claim["derivation"]
    require(isinstance(definition, dict) and set(definition) == {"operation", "inputs", "rounding"},
            f"{identifier}: derivation needs operation, inputs, and rounding")
    inputs = definition["inputs"]
    operation = definition["operation"]
    require(isinstance(inputs, list) and inputs and all(text(ref) for ref in inputs),
            f"{identifier}: derivation inputs must be claim IDs")
    require(operation in ("sum", "difference", "product", "ratio")
            and (operation in ("sum", "product") or len(inputs) == 2),
            f"{identifier}: unsupported operation or input arity")
    rounding = definition["rounding"]
    require(integer(rounding) and 0 <= rounding <= 12,
            f"{identifier}: rounding must be an integer from 0 to 12")
    require(text(claim.get("units")), f"{identifier}: derived claim requires units")
    visiting = visiting | {identifier}
    values = []
    for ref in inputs:
        value = evaluate_claim(ref, claims, visiting)
        require(type(value) in (int, float, str) and type(value) is not bool,
                f"{identifier}: derivation input must be numeric")
        try:
            number = Decimal(str(value))
        except InvalidOperation as exc:
            raise Invalid(f"{identifier}: nonnumeric derivation input") from exc
        require(number.is_finite(), f"{identifier}: nonfinite input")
        values.append(number)
    input_units = [claims[ref].get("units") for ref in inputs]
    require(all(text(unit) for unit in input_units), f"{identifier}: input units required")
    if operation in ("sum", "difference"):
        require(all(unit == claim["units"] for unit in input_units),
                f"{identifier}: incompatible input units")
    if operation == "sum":
        result = sum(values, Decimal(0))
    elif operation == "difference":
        result = values[0] - values[1]
    elif operation == "product":
        result = Decimal(1)
        for value in values:
            result *= value
    else:
        require(values[1] != 0, f"{identifier}: division by zero")
        require(input_units[0] == input_units[1] and claim["units"] in ("ratio", "percent"),
                f"{identifier}: ratio needs equal input units and ratio/percent output")
        result = values[0] / values[1]
        if claim["units"] == "percent":
            result *= 100
    try:
        result = result.quantize(Decimal(1).scaleb(-rounding), rounding=ROUND_HALF_UP)
        supplied = Decimal(str(claim.get("value")))
    except InvalidOperation as exc:
        raise Invalid(f"{identifier}: invalid derived value or precision") from exc
    require(supplied.is_finite() and supplied == result, f"{identifier}: incorrect derived value")
    return format(result, "f")


def check_draft(canon, plan, draft):
    documents = validate_plan(canon, plan)
    index = {record["id"]: record for record in canon["records"]}
    _check_draft(documents, index, fingerprint(canon), fingerprint(plan), draft, canon["corpus_id"])


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


def _check_draft(documents, index, canon_sha256, plan_sha256, draft, corpus_id):
    version(draft, "Draft")
    require(draft.get("corpus_id") == corpus_id, "Draft corpus identity mismatch")
    identifier = draft.get("document_id")
    require(isinstance(identifier, str) and identifier in documents, f"Unknown document: {identifier}")
    doc = documents[identifier]
    for key, value in (("canon_sha256", canon_sha256), ("plan_sha256", plan_sha256),
                       ("title", doc["title"]),
                       ("content_depth", doc["content_depth"]),
                       ("publication_date", doc["publication_date"])):
        require(type(draft.get(key)) is type(value) and draft.get(key) == value,
                f"{doc['id']}: wrong draft {key}")
    visible = project_records(index, doc, for_claims=True)
    recipients = (["public"] if doc["audience"] == "public" else [
        item["id"] for item in index.values() if item["kind"] == "person"
        and (not doc["access_groups"] or
             principal_groups(index, item["id"], doc["publication_date"]) & set(doc["access_groups"]))])
    claims = draft.get("claims")
    require(isinstance(claims, list) and claims, f"{doc['id']}: claims required")
    by_id, cited_records = {}, set()
    for claim in claims:
        require(isinstance(claim, dict) and text(claim.get("id")), "Invalid claim")
        identifier = claim["id"]
        require(identifier not in by_id, f"Duplicate claim ID: {identifier}")
        if "derivation" in claim:
            continue
        ref = claim.get("record_id")
        require(isinstance(ref, str) and ref in visible,
                f"{identifier}: unknown or future record {ref}, or missing from the bounded canonical org graph")
        require(all(accessible(visible[ref], index, person, doc["publication_date"])
                    for person in recipients), f"{identifier}: not released to the intended audience")
        expected = resolve(visible[ref], claim.get("path"))
        require(fingerprint(expected) == fingerprint(resolve(
            writer_record(visible[ref], visible), claim.get("path"))),
            f"{identifier}: private policy references cannot be claims")
        require("value" in claim and fingerprint(claim["value"]) == fingerprint(expected),
                f"{identifier}: value differs from canon")
        by_id[identifier] = claim
        cited_records.add(ref)
        statement_kind = visible[ref]["data"]["kind"] if visible[ref]["kind"] == "statement" else "fact"
        require(claim.get("statement_kind", statement_kind) == statement_kind,
                f"{identifier}: claim misclassifies its statement kind")
        if "units" in claim:
            require(text(claim["units"]), f"{identifier}: units must be nonempty text")
            if claim["path"] == ["data", "value"] and "units" in visible[ref]["data"]:
                require(claim["units"] == visible[ref]["data"]["units"],
                        f"{identifier}: units differ from the source")
            if len(claim["path"]) == 2 and claim["path"][0] == "data" and isinstance(
                    claim["path"][1], str) and claim["path"][1].endswith("_usd"):
                require(claim["units"] == "USD", f"{identifier}: USD actuals require USD units")
    for claim in claims:
        if "derivation" in claim:
            identifier = claim["id"]
            require(identifier not in by_id, f"Duplicate claim ID: {identifier}")
            by_id[identifier] = claim
    for identifier in by_id:
        evaluate_claim(identifier, by_id)
    require(set(doc["anchor_ids"]) <= cited_records, f"{doc['id']}: uncited anchors")
    represented = {claim["record_id"] for claim in claims if "record_id" in claim
                   and claim.get("path") == ["data", "value"]}
    require(set(doc["required_statement_ids"]) <= represented,
            f"{doc['id']}: required source assertions must be represented as value claims")
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
        if "speaker_id" in section:
            speaker = section["speaker_id"]
            require(speaker in visible and visible[speaker]["kind"] == "person",
                    f"{doc['id']}: unknown or hidden section speaker")
            require(type(section.get("reported_speech", False)) is bool,
                    f"{doc['id']}: reported_speech must be boolean")
            if not section.get("reported_speech", False):
                require(all("record_id" not in by_id[ref]
                            or visible[by_id[ref]["record_id"]]["kind"] != "statement"
                            or visible[by_id[ref]["record_id"]]["data"]["speaker_id"] == speaker
                            for ref in refs), f"{doc['id']}: statement attributed to the wrong speaker")
    require(used == by_id.keys(), f"{doc['id']}: unused claims")
    require(set(doc["required_sections"]) <= roles,
            f"{doc['id']}: missing genre sections {set(doc['required_sections']) - roles}")


def check_drafts(canon, plan, directory):
    documents = validate_plan(canon, plan)
    index = {record["id"]: record for record in canon["records"]}
    canon_sha256, plan_sha256 = fingerprint(canon), fingerprint(plan)
    files = list(Path(directory).glob("*.json"))
    require({p.stem for p in files} == documents.keys(), "Draft IDs do not exactly match plan")
    for path in files:
        draft = load(path)
        require(isinstance(draft, dict) and draft.get("document_id") == path.stem,
                f"{path}: filename does not match document_id")
        _check_draft(documents, index, canon_sha256, plan_sha256, draft, canon["corpus_id"])
    return len(files)


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def normalized(value):
    return " ".join(str(value).split())


def xml_part(package, part):
    require(part in package.namelist(), f"Missing related Office part: {part}")
    return ElementTree.fromstring(package.read(part))


def relationships(package, part):
    folder, name = posixpath.split(part)
    rel_part = posixpath.join(folder, "_rels", name + ".rels")
    if rel_part not in package.namelist():
        return {}
    result = {}
    for relation in xml_part(package, rel_part):
        identifier = relation.attrib.get("Id")
        require(text(identifier) and identifier not in result, f"{part}: duplicate relationship ID")
        target = relation.attrib.get("Target", "")
        external = relation.attrib.get("TargetMode") == "External"
        resolved = posixpath.normpath(target.lstrip("/") if target.startswith("/")
                                     else posixpath.join(folder, target))
        require(external or (not resolved.startswith("../") and resolved != ".."
                             and "\\" not in resolved and ":" not in resolved),
                f"{part}: escaping relationship target")
        result[identifier] = {"part": resolved, "type": relation.attrib.get("Type", ""),
                              "external": external}
    return result


def related_part(rels, identifier, expected):
    require(identifier in rels and not rels[identifier]["external"]
            and rels[identifier]["type"].endswith("/" + expected),
            f"Missing, external, or wrong {expected} relationship: {identifier}")
    return rels[identifier]["part"]


def element_text(element, namespace):
    pieces = []
    for item in element.iter():
        local = item.tag.rsplit("}", 1)[-1]
        if item.tag == "{" + NS[namespace] + "}t":
            pieces.append(item.text or "")
        elif local in ("tab", "br", "cr"):
            pieces.append(" ")
    return normalized("".join(pieces))


def joined_container(locator, fragments, **metadata):
    content, segments = "", []
    for unit, value in fragments:
        value = normalized(value)
        if not value:
            continue
        start = len(content) + bool(content)
        content += (" " if content else "") + value
        segments.append({"unit": unit, "start": start, "end": len(content)})
    return dict(metadata, locator=locator, text=content, segments=segments)


def chart_containers(package, owner_part, owner, context):
    entries = []
    rels = relationships(package, owner_part)
    for chart in owner.findall(".//c:chart", NS):
        part = related_part(rels, chart.get("{" + NS["r"] + "}id"), "chart")
        tree = xml_part(package, part)
        for series_index, series in enumerate(tree.findall(".//c:ser", NS)):
            for section in ("tx", "cat", "val", "xVal", "yVal", "bubbleSize"):
                node = series.find("c:" + section, NS)
                if node is None:
                    continue
                formula = node.find(".//c:f", NS)
                for point in node.findall(".//c:pt", NS):
                    value = point.find("c:v", NS)
                    if value is None:
                        continue
                    locator = dict(context, part=part, kind="chart_point", series=series_index,
                                   component=section, point=point.get("idx"))
                    entries.append({"locator": locator, "text": normalized(value.text or ""),
                                    "formula": formula.text if formula is not None else None})
    return entries


def word_containers(package):
    main = "word/document.xml"
    tree = xml_part(package, main)
    parts = [main]
    rels = relationships(package, main)
    for key in ("header", "footer"):
        for reference in tree.findall(".//w:" + key + "Reference", NS):
            parts.append(related_part(rels, reference.get("{" + NS["r"] + "}id"), key))
    entries = []
    for part in dict.fromkeys(parts):
        story = xml_part(package, part)
        paragraphs = story.findall(".//w:p", NS)
        paragraph_ids = {paragraph: number for number, paragraph in enumerate(paragraphs)}
        def fragments(items):
            return [({"part": part, "paragraph": paragraph_ids[item]}, element_text(item, "w"))
                    for item in items]
        for number, paragraph in enumerate(paragraphs):
            entries.append(joined_container(
                {"kind": "paragraph", "part": part, "index": number}, fragments([paragraph])))
            entries.extend(chart_containers(package, part, paragraph, {"paragraph": number}))
        for table_index, table in enumerate(story.findall(".//w:tbl", NS)):
            for row_index, row in enumerate(table.findall("w:tr", NS)):
                for column, cell in enumerate(row.findall("w:tc", NS)):
                    entries.append(joined_container(
                        {"kind": "table_cell", "part": part, "table": table_index,
                         "row": row_index, "column": column}, fragments(cell.findall(".//w:p", NS))))
    return entries


def slide_shapes(package, part, slide_id, notes=False):
    tree = xml_part(package, part)
    entries, seen = [], set()
    paragraph_ids = {paragraph: number for number, paragraph in enumerate(tree.findall(".//a:p", NS))}
    def fragments(items):
        return [({"part": part, "paragraph": paragraph_ids[item]}, element_text(item, "a"))
                for item in items]
    for shape in tree.findall(".//p:sp", NS) + tree.findall(".//p:graphicFrame", NS):
        props = shape.find(".//p:cNvPr", NS)
        require(props is not None and text(props.get("id")), f"{part}: shape ID missing")
        identifier = props.get("id")
        require(identifier not in seen, f"{part}: duplicate shape ID")
        seen.add(identifier)
        context = {"slide_id": slide_id, "shape_id": identifier}
        entries.append(joined_container(
            dict(context, part=part, kind="notes_shape" if notes else "shape"),
            fragments(shape.findall(".//a:p", NS))))
        for row_index, row in enumerate(shape.findall(".//a:tbl/a:tr", NS)):
            for column, cell in enumerate(row.findall("a:tc", NS)):
                entries.append(joined_container(
                    dict(context, part=part, kind="slide_cell", row=row_index, column=column),
                    fragments(cell.findall(".//a:p", NS))))
        entries.extend(chart_containers(package, part, shape, context))
    return entries


def presentation_containers(package):
    main = "ppt/presentation.xml"
    tree = xml_part(package, main)
    rels = relationships(package, main)
    entries, seen = [], set()
    for slide in tree.findall("p:sldIdLst/p:sldId", NS):
        identifier = slide.get("id")
        require(text(identifier) and identifier not in seen, "Duplicate or missing slide ID")
        seen.add(identifier)
        part = related_part(rels, slide.get("{" + NS["r"] + "}id"), "slide")
        entries.extend(slide_shapes(package, part, identifier))
        slide_rels = relationships(package, part)
        for ref, relation in slide_rels.items():
            if relation["type"].endswith("/notesSlide"):
                entries.extend(slide_shapes(package, related_part(slide_rels, ref, "notesSlide"),
                                             identifier, notes=True))
    require(seen, "Presentation has no relationship-ordered slides")
    return entries


def cell_coordinates(address):
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]*)", address)
    require(match is not None, f"Invalid cell address: {address}")
    column = 0
    for char in match[1]:
        column = column * 26 + ord(char) - ord("A") + 1
    row = int(match[2])
    require(column <= 16384 and row <= 1048576, f"Cell outside worksheet bounds: {address}")
    return row, column


def spreadsheet_containers(package):
    main = "xl/workbook.xml"
    tree = xml_part(package, main)
    rels = relationships(package, main)
    strings = []
    for ref, relation in rels.items():
        if relation["type"].endswith("/sharedStrings"):
            part = related_part(rels, ref, "sharedStrings")
            strings = [element_text(item, "x") for item in xml_part(package, part).findall("x:si", NS)]
    entries, seen = [], set()
    for sheet in tree.findall("x:sheets/x:sheet", NS):
        name, sheet_id = sheet.get("name"), sheet.get("sheetId")
        require(text(name) and name not in seen and text(sheet_id), "Invalid or duplicate worksheet")
        seen.add(name)
        part = related_part(rels, sheet.get("{" + NS["r"] + "}id"), "worksheet")
        worksheet = xml_part(package, part)
        shared_formulas = {}
        for cell in worksheet.findall(".//x:sheetData/x:row/x:c", NS):
            formula = cell.find("x:f", NS)
            if formula is not None and formula.get("t") == "shared" and text(formula.text):
                identifier = formula.get("si")
                addresses = formula.get("ref", "").split(":")
                require(text(identifier) and identifier not in shared_formulas and len(addresses) in (1, 2),
                        f"{name}: invalid shared formula base")
                start, end = cell_coordinates(addresses[0]), cell_coordinates(addresses[-1])
                require(start[0] <= end[0] and start[1] <= end[1],
                        f"{name}: reversed shared formula range")
                shared_formulas[identifier] = {
                    "kind": "shared", "index": identifier, "base_cell": cell.get("r"),
                    "base_formula": formula.text, "range": formula.get("ref")}
        cells = set()
        for cell in worksheet.findall(".//x:sheetData/x:row/x:c", NS):
            address = cell.get("r", "")
            cell_coordinates(address)
            require(address not in cells, f"{name}: duplicate cell {address}")
            cells.add(address)
            kind = cell.get("t")
            value, formula = cell.find("x:v", NS), cell.find("x:f", NS)
            raw = value.text or "" if value is not None else ""
            if kind == "s":
                require(re.fullmatch(r"\d+", raw) is not None and int(raw) < len(strings),
                        f"{name}!{address}: invalid shared string")
                content = strings[int(raw)]
            elif kind == "inlineStr":
                content = element_text(cell, "x")
            elif kind == "b":
                require(raw in ("0", "1"), f"{name}!{address}: invalid boolean")
                content = "true" if raw == "1" else "false"
            else:
                content = normalized(raw)
            require(kind != "e", f"{name}!{address}: Excel error {raw}")
            require(formula is None or value is not None,
                    f"{name}!{address}: formula has no saved cached value; recalculate in Excel")
            saved_formula = formula.text if formula is not None else None
            if formula is not None and formula.get("t") == "shared":
                identifier = formula.get("si")
                require(identifier in shared_formulas, f"{name}!{address}: shared formula base missing")
                definition = shared_formulas[identifier]
                addresses = definition["range"].split(":")
                start, end = cell_coordinates(addresses[0]), cell_coordinates(addresses[-1])
                row, column = cell_coordinates(address)
                require(start[0] <= row <= end[0] and start[1] <= column <= end[1],
                        f"{name}!{address}: cell outside its shared formula range")
                if not text(saved_formula):
                    saved_formula = dict(definition, cell=address)
            entries.append({"locator": {"kind": "cell", "part": part, "sheet_id": sheet_id,
                                        "sheet": name, "cell": address},
                            "text": content, "formula": saved_formula})
        sheet_rels = relationships(package, part)
        for drawing in worksheet.findall("x:drawing", NS):
            drawing_part = related_part(sheet_rels, drawing.get("{" + NS["r"] + "}id"), "drawing")
            entries.extend(chart_containers(package, drawing_part, xml_part(package, drawing_part),
                                             {"sheet_id": sheet_id, "sheet": name}))
    require(seen, "Workbook has no relationship-ordered worksheets")
    return entries


def artifact_containers(path, fmt):
    audit_package(path, fmt)
    try:
        with zipfile.ZipFile(path) as package:
            result = {"docx": word_containers, "pptx": presentation_containers,
                      "xlsx": spreadsheet_containers}[fmt](package)
    except (zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise Invalid(f"{path}: invalid Office package: {exc}") from exc
    keys = [fingerprint(item["locator"]) for item in result]
    require(len(keys) == len(set(keys)), f"{path}: ambiguous duplicate artifact locators")
    return result


def manifest_files(canon, plan, manifest, root):
    documents = validate_plan(canon, plan)
    version(manifest, "Manifest")
    require(manifest.get("corpus_id") == canon["corpus_id"]
            and manifest.get("canon_sha256") == fingerprint(canon)
            and manifest.get("plan_sha256") == fingerprint(plan), "Manifest fingerprint/version mismatch")
    root = Path(root).resolve()
    validate_output_layout(canon, root)
    entries = manifest.get("documents")
    require(isinstance(entries, list), "Manifest documents must be a list")
    files, paths = {}, set()
    for entry in entries:
        require(isinstance(entry, dict), "Invalid manifest document")
        identifier = entry.get("document_id")
        require(identifier in documents and identifier not in files,
                f"Invalid/duplicate manifest document: {identifier}")
        relative = entry.get("path")
        require(text(relative) and not Path(relative).is_absolute(), f"{identifier}: path must be relative")
        path = (root / relative).resolve()
        require(path.is_relative_to(root) and path not in paths,
                f"{identifier}: duplicate or escaping artifact path")
        require(path.is_relative_to(root / "documents"), f"{identifier}: final artifacts must be in documents")
        require(path.is_file(), f"{identifier}: missing artifact {path}")
        doc = documents[identifier]
        require(path.suffix == "." + doc["format"], f"{identifier}: wrong extension")
        require(path.name == doc["filename"], f"{identifier}: artifact filename does not match the descriptive planned filename")
        require(entry.get("sha256") == file_hash(path), f"{identifier}: artifact hash mismatch")
        files[identifier] = path
        paths.add(path)
    require(files.keys() == documents.keys(), "Manifest is incomplete")
    require({p.resolve() for p in (root / "documents").rglob("*") if p.is_file()} == paths,
            "documents must contain only the manifested final artifacts; put other files in build")
    return files


def receipt_header(canon, plan):
    return {"schema_version": SCHEMA_VERSION, "corpus_id": canon["corpus_id"],
            "canon_sha256": fingerprint(canon), "plan_sha256": fingerprint(plan)}


def index_artifacts(canon, plan, manifest, root):
    files = manifest_files(canon, plan, manifest, root)
    return dict(receipt_header(canon, plan), documents=[
        {"document_id": doc["id"], "revision": doc["revision"],
         "artifact_sha256": file_hash(files[doc["id"]]),
         "containers": artifact_containers(files[doc["id"]], doc["format"])}
        for doc in plan["documents"]])


def display_value(claim, rendering=None):
    rendering = {} if rendering is None else rendering
    require(isinstance(rendering, dict) and not set(rendering) - {
        "scale", "decimals", "grouping", "prefix", "suffix"}, "Unknown rendering fields")
    value = claim["value"]
    if not rendering:
        return normalized(str(value).lower() if type(value) is bool else value)
    require(type(value) in (int, float, str), "Rendering requires a numeric scalar claim")
    try:
        number, scale = Decimal(str(value)), Decimal(str(rendering.get("scale", 1)))
    except InvalidOperation as exc:
        raise Invalid("Rendering value and scale must be numeric") from exc
    require(number.is_finite() and scale.is_finite() and scale > 0, "Invalid rendering scale/value")
    decimals = rendering.get("decimals", 0)
    require(integer(decimals) and 0 <= decimals <= 12, "Rendering decimals must be from 0 to 12")
    require(type(rendering.get("grouping", False)) is bool, "Rendering grouping must be boolean")
    require(all(isinstance(rendering.get(key, ""), str) for key in ("prefix", "suffix")),
            "Rendering prefix/suffix must be strings")
    number = (number / scale).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    spec = ("," if rendering.get("grouping") else "") + f".{decimals}f"
    return normalized(rendering.get("prefix", "") + format(number, spec) + rendering.get("suffix", ""))


def located_container(containers, locator):
    require(isinstance(locator, dict), "Locator must be an object")
    if locator.get("kind") != "range":
        matches = [entry for entry in containers if entry["locator"] == locator]
        require(len(matches) == 1, "Locator missing or ambiguous; specify an exact indexed container")
        return matches[0]
    require(set(locator) == {"kind", "part", "sheet_id", "sheet", "range"}, "Invalid range locator")
    addresses = locator["range"].split(":") if isinstance(locator.get("range"), str) else []
    require(len(addresses) == 2, "Range must have start:end addresses")
    start, end = map(cell_coordinates, addresses)
    require(start[0] <= end[0] and start[1] <= end[1], "Reversed cell range")
    cells = [entry for entry in containers if entry["locator"]["kind"] == "cell"
             and all(entry["locator"].get(key) == locator[key] for key in ("part", "sheet_id", "sheet"))
             and start[0] <= cell_coordinates(entry["locator"]["cell"])[0] <= end[0]
             and start[1] <= cell_coordinates(entry["locator"]["cell"])[1] <= end[1]]
    require(len(cells) == (end[0] - start[0] + 1) * (end[1] - start[1] + 1),
            "Range contains missing/unindexed cells")
    cells.sort(key=lambda entry: cell_coordinates(entry["locator"]["cell"]))
    return joined_container(locator, [(entry["locator"], entry["text"]) for entry in cells],
                            formula=[entry.get("formula") for entry in cells])


def verify_provenance(canon, plan, drafts, manifest, root, locations):
    check_drafts(canon, plan, drafts)
    version(locations, "Claim locations")
    header = receipt_header(canon, plan)
    require(all(locations.get(key) == value for key, value in header.items()),
            "Claim locations fingerprint/identity mismatch")
    artifact_index = index_artifacts(canon, plan, manifest, root)
    indexed = {doc["document_id"]: doc for doc in artifact_index["documents"]}
    documents = {doc["id"]: doc for doc in plan["documents"]}
    records = {record["id"]: record for record in canon["records"]}
    mappings = locations.get("documents")
    require(isinstance(mappings, list), "Claim locations documents must be a list")
    result, seen = [], set()
    for mapping in mappings:
        require(isinstance(mapping, dict) and set(mapping) == {
            "document_id", "revision", "artifact_sha256", "draft_sha256", "occurrences"},
            "Invalid location document fields")
        identifier = mapping.get("document_id")
        require(identifier in documents and identifier not in seen, "Duplicate/unknown location document")
        seen.add(identifier)
        doc, indexed_doc = documents[identifier], indexed[identifier]
        require(mapping["revision"] == doc["revision"], f"{identifier}: wrong mapping revision")
        draft_path = Path(drafts) / f"{identifier}.json"
        draft = load(draft_path)
        require(mapping.get("artifact_sha256") == indexed_doc["artifact_sha256"]
                and mapping.get("draft_sha256") == file_hash(draft_path),
                f"{identifier}: stale artifact/draft location mapping")
        claims = {claim["id"]: claim for claim in draft["claims"]}
        occurrences = mapping.get("occurrences")
        require(isinstance(occurrences, list) and occurrences, f"{identifier}: occurrences required")
        represented, verified, occurrence_keys, physical_spans = set(), [], set(), {}
        for occurrence in occurrences:
            require(isinstance(occurrence, dict) and not set(occurrence) - {
                "claim_id", "section_index", "locator", "start", "end", "expected_text",
                "rendering", "formula"}, "Invalid occurrence fields")
            claim_id, section = occurrence.get("claim_id"), occurrence.get("section_index")
            require(claim_id in claims and integer(section) and 0 <= section < len(draft["sections"])
                    and claim_id in draft["sections"][section]["claim_ids"],
                    f"{identifier}: occurrence must reference its draft section and claim")
            claim = claims[claim_id]
            container = located_container(indexed_doc["containers"], occurrence.get("locator"))
            start, end = occurrence.get("start"), occurrence.get("end")
            require(integer(start) and integer(end) and 0 <= start < end <= len(container["text"]),
                    f"{identifier}: occurrence span outside normalized container")
            excerpt = container["text"][start:end]
            require(occurrence.get("expected_text") == excerpt
                    and display_value(claim, occurrence.get("rendering")) == excerpt,
                    f"{identifier}: located text does not match its grounded claim/rendering")
            if "formula" in occurrence:
                require(occurrence["formula"] == container.get("formula"),
                        f"{identifier}: saved formula differs from occurrence")
            if "derivation" in claim and occurrence["locator"]["kind"] in ("cell", "range"):
                formulas = container.get("formula")
                formulas = formulas if isinstance(formulas, list) else [formulas]
                require("formula" in occurrence and formulas and all(
                    text(value) or isinstance(value, dict) for value in formulas),
                        f"{identifier}: derived cell/range claims require saved formula evidence")
            key = (fingerprint(occurrence["locator"]), start, end, claim_id)
            require(key not in occurrence_keys, f"{identifier}: duplicate occurrence")
            occurrence_keys.add(key)
            segments = container.get("segments", [
                {"unit": container["locator"], "start": 0, "end": len(container["text"])}])
            for segment in segments:
                left, right = max(start, segment["start"]), min(end, segment["end"])
                if left >= right:
                    continue
                span = (left - segment["start"], right - segment["start"])
                unit = (claim_id, fingerprint(segment["unit"]))
                require(all(span[0] >= previous[1] or span[1] <= previous[0]
                            for previous in physical_spans.get(unit, [])),
                        f"{identifier}: duplicate occurrence through overlapping container aliases")
                physical_spans.setdefault(unit, []).append(span)
            item = dict(occurrence)
            if "record_id" in claim:
                record = records[claim["record_id"]]
                item["record_id"] = record["id"]
                if record["kind"] == "statement":
                    item["statement_kind"] = record["data"]["kind"]
                    item["source_version_id"] = record["data"]["source_version_id"]
                    item["speaker_id"] = record["data"]["speaker_id"]
                    item["qualifiers"] = record["data"].get("qualifiers", [])
            else:
                item["input_claim_ids"] = claim["derivation"]["inputs"]
            verified.append(item)
            represented.add((section, claim_id))
        expected_occurrences = {(number, ref) for number, section in enumerate(draft["sections"])
                                for ref in section["claim_ids"]}
        require(represented == expected_occurrences, f"{identifier}: final-artifact claim coverage incomplete")
        result.append({"document_id": identifier, "revision": doc["revision"],
                       "artifact_sha256": indexed_doc["artifact_sha256"],
                       "draft_sha256": file_hash(draft_path), "occurrences": verified})
    require(seen == documents.keys(), "Final-artifact provenance documents incomplete")
    result.sort(key=lambda item: item["document_id"])
    return dict(header, documents=result,
                validation="identity, exact occurrence content, and grounded receipts; semantic QA required")


def validate_bindings(canon, plan, manifest, bindings):
    version(bindings, "Resource bindings")
    require(bindings.get("corpus_id") == canon["corpus_id"], "Binding corpus identity mismatch")
    docs = {doc["id"]: doc for doc in plan["documents"]}
    hashes = {entry["document_id"]: entry["sha256"] for entry in manifest["documents"]}
    entries = bindings.get("bindings")
    require(isinstance(entries, list), "bindings must be a list")
    seen = set()
    for entry in entries:
        require(isinstance(entry, dict) and not set(entry) - {
            "document_id", "revision", "artifact_sha256", "kind", "path", "drive_id", "item_id"},
            "Unknown resource binding fields")
        identifier, kind = entry.get("document_id"), entry.get("kind")
        require(identifier in docs and kind in ("local", "sharepoint"), "Unknown binding document/kind")
        require(entry.get("revision") == docs[identifier]["revision"]
                and entry.get("artifact_sha256") == hashes[identifier], "Stale binding revision/hash")
        require((identifier, kind) not in seen, "Duplicate binding kind for document")
        seen.add((identifier, kind))
        if kind == "local":
            require(text(entry.get("path")) and not Path(entry["path"]).is_absolute()
                    and ".." not in entry["path"].replace("\\", "/").split("/"),
                    "Local binding requires a safe relative path")
            require("drive_id" not in entry and "item_id" not in entry, "Mixed binding kinds")
        else:
            require(text(entry.get("drive_id")) and text(entry.get("item_id")) and "path" not in entry,
                    "SharePoint binding requires drive_id and item_id, not a path")
    return len(entries)


def coverage_report(canon, plan, drafts, provenance):
    header = receipt_header(canon, plan)
    require(all(provenance.get(key) == value for key, value in header.items()),
            "Provenance coverage fingerprint mismatch")
    documents = validate_plan(canon, plan)
    delivered = {item["document_id"]: item for item in provenance["documents"]}
    require(delivered.keys() == documents.keys(), "Coverage provenance incomplete")
    index = {record["id"]: record for record in canon["records"]}
    used, statement_kinds, qualifiers = set(), Counter(), Counter()
    for identifier, mapped in delivered.items():
        draft = load(Path(drafts) / f"{identifier}.json")
        require(mapped["draft_sha256"] == file_hash(Path(drafts) / f"{identifier}.json"),
                "Stale coverage draft")
        used.update(claim["record_id"] for claim in draft["claims"] if "record_id" in claim)
    for ref in used:
        if index[ref]["kind"] == "statement":
            statement_kinds[index[ref]["data"]["kind"]] += 1
            qualifiers.update(index[ref]["data"].get("qualifiers", []))
    instances = condition_instances(canon)
    conditions = []
    for condition in instances:
        data, kind = condition["data"], condition["data"]["kind"]
        carriers = [doc for doc in plan["documents"] if condition["id"] in doc["condition_ids"]]
        require(carriers, f"{condition['id']}: no delivered carrier")
        if kind == "conflict":
            require(set(data["statement_ids"]) <= used, f"{condition['id']}: conflict assertions not delivered")
        if kind in ("unknown", "undocumented"):
            for doc in carriers:
                draft = load(Path(drafts) / f"{doc['id']}.json")
                refs = {claim["record_id"] for claim in draft["claims"] if "record_id" in claim}
                require(not refs & set(data["blocked_record_ids"]), f"{condition['id']}: gap was filled")
        conditions.append({"condition_id": condition["id"], "request_id": data["request_id"],
                           "kind": kind, "documents": [doc["id"] for doc in carriers],
                           "validation": "structured references; semantic condition review required"})
    def counts(key):
        return dict(sorted(Counter(doc[key] for doc in plan["documents"]).items()))
    perspectives = Counter(doc["viewpoint_id"] for doc in plan["documents"] if doc["viewpoint_id"])
    publication_years = dict(sorted(Counter(doc["publication_date"][:4]
                                           for doc in plan["documents"]).items()))
    topic_years = dict(sorted(Counter(str(topic_year(index[doc["anchor_ids"][0]]))
                                     for doc in plan["documents"]).items()))
    periods = Counter()
    for doc in plan["documents"]:
        periods.update(document_periods(doc, index))
    periods = dict(sorted(periods.items()))
    planned_sources = sorted({doc["source_version_id"] for doc in plan["documents"]
                              if doc["source_version_id"]})
    return dict(header, planned={"documents": len(documents), "formats": counts("format"),
                                "types": counts("type"), "viewpoints": dict(sorted(perspectives.items())),
                                "publication_years": publication_years, "topic_years": topic_years,
                                "periods": periods, "source_versions": planned_sources},
                delivered={"documents": len(delivered), "formats": counts("format"),
                           "types": counts("type"), "viewpoints": dict(sorted(perspectives.items())),
                           "publication_years": publication_years, "topic_years": topic_years,
                           "periods": periods,
                           "record_ids": sorted(used),
                           "topics": sorted({ref for doc in plan["documents"] for ref in doc["anchor_ids"]}),
                           "entities": sorted(ref for ref in used if index[ref]["kind"] in (
                               "company", "person", "department", "product")),
                           "subjects": sorted({index[ref]["data"]["subject_id"] for ref in used
                                               if index[ref]["kind"] == "statement"}),
                           "statement_kinds": dict(sorted(statement_kinds.items())),
                           "qualifiers": dict(sorted(qualifiers.items())),
                           "source_versions": sorted(set(planned_sources) | {
                               index[ref]["data"]["source_version_id"] for ref in used
                               if index[ref]["kind"] == "statement"}),
                           "claims": sum(len(load(Path(drafts) / f"{ref}.json")["claims"])
                                         for ref in delivered),
                           "occurrences": sum(len(item["occurrences"]) for item in delivered.values()),
                           "unverified_claims": 0},
                conditions=conditions,
                denominator="requested documents, configured conditions, and explicitly grounded claims",
                validation="delivered structural coverage, not semantic depth or agent performance")


def access_view(canon, plan, manifest, provenance, principal, as_of):
    iso(as_of)
    documents = validate_plan(canon, plan)
    index = {record["id"]: record for record in canon["records"]}
    require(principal == "public" or (principal in index and index[principal]["kind"] == "person"
            and index[principal]["known_on"] <= as_of), "Unknown or future access-view principal")
    require(all(provenance.get(key) == value for key, value in receipt_header(canon, plan).items()),
            "Access-view provenance fingerprint mismatch")
    hashes = {entry["document_id"]: entry["sha256"] for entry in manifest["documents"]}
    mapped = {item["document_id"]: item for item in provenance["documents"]}
    groups = principal_groups(index, principal, as_of)
    selected = []
    for identifier, doc in documents.items():
        if doc["publication_date"] > as_of or (doc["access_groups"] and not groups & set(doc["access_groups"])):
            continue
        if principal == "public" and doc["audience"] != "public":
            continue
        require(identifier in mapped and mapped[identifier]["artifact_sha256"] == hashes[identifier],
                "Access-view artifact provenance mismatch")
        visible = project_records(index, dict(doc, author_id=principal, publication_date=as_of,
                                              blocked_record_ids=[], suppressed_record_ids=[]),
                                  include_knowledge=False)
        occurrences = []
        for occurrence in mapped[identifier]["occurrences"]:
            item = {key: value for key, value in occurrence.items() if key in (
                "claim_id", "section_index", "locator", "start", "end", "expected_text", "rendering", "formula")}
            for key in ("record_id", "source_version_id", "speaker_id"):
                if occurrence.get(key) in visible:
                    item[key] = occurrence[key]
            occurrences.append(item)
        selected.append({"document_id": identifier, "revision": doc["revision"],
                         "artifact_sha256": hashes[identifier], "filename": doc["filename"],
                         "occurrences": occurrences})
    return {"schema_version": SCHEMA_VERSION, "corpus_id": canon["corpus_id"],
            "principal_id": principal, "as_of_date": as_of, "documents": selected,
            "validation": "simulated selection only; not ACL enforcement or secure distribution"}


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
    require_build_path(drafts, root)
    files = manifest_files(canon, plan, manifest, root)
    hashes = {identifier: file_hash(path) for identifier, path in files.items()}
    for identifier, path in files.items():
        audit_package(path, documents[identifier]["format"])
    qa = load(root / "build" / "qa.json")
    require(isinstance(qa, dict) and qa.get("schema_version") == SCHEMA_VERSION
            and qa.get("corpus_id") == canon["corpus_id"]
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
        if documents[identifier]["condition_ids"]:
            gates.append("condition_review")
        for gate in gates:
            receipt = review.get(gate)
            require(isinstance(receipt, dict) and receipt.get("status") == "passed"
                    and text(receipt.get("evidence")), f"{identifier}: missing/failed {gate}")
            if gate == "depth_review":
                require(receipt.get("content_depth") == documents[identifier]["content_depth"],
                        f"{identifier}: depth_review content_depth mismatch")
        reviewed.add(identifier)
    require(reviewed == documents.keys(), "QA receipts incomplete")
    provenance = verify_provenance(canon, plan, drafts, manifest, root,
                                  load(root / "build" / "claim-locations.json"))
    require(provenance == load(root / "build" / "provenance.json"), "Stale or forged provenance receipt")
    require(coverage_report(canon, plan, drafts, provenance) == load(root / "build" / "coverage.json"),
            "Stale or incorrect coverage receipt")
    binding_path = root / "build" / "resource-bindings.json"
    if binding_path.exists():
        validate_bindings(canon, plan, manifest, load(binding_path))
    return {"documents": len(files), "formats": {
        fmt: sum(d["format"] == fmt for d in documents.values()) for fmt in FORMATS
    }, "validation": "structure, hashes, claims, final-artifact provenance, coverage, and QA receipts; "
                      "not independent semantic certification"}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="Create a seeded baseline canon in a new directory")
    init.add_argument("--config", required=True)
    init.add_argument("--out", required=True)
    artifact_commands = ("index-artifacts", "verify-provenance", "coverage", "access-view")
    for name in ("validate", "plan", "context", "check", "audit") + artifact_commands:
        command = commands.add_parser(name)
        command.add_argument("--canon", required=True)
        if name in ("context", "check", "audit") + artifact_commands:
            command.add_argument("--plan", required=True)
        if name in ("plan", "context") + artifact_commands:
            command.add_argument("--out", required=True)
        if name == "context":
            command.add_argument("--doc", required=True)
        if name in ("check", "audit", "verify-provenance", "coverage", "access-view"):
            command.add_argument("--drafts", required=True)
        if name in ("audit",) + artifact_commands:
            command.add_argument("--manifest", required=True)
            command.add_argument("--root", required=True)
        if name in ("verify-provenance", "coverage"):
            command.add_argument("--locations", required=True)
        if name == "access-view":
            command.add_argument("--principal", required=True)
            command.add_argument("--as-of", required=True)
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
            for key in ("out", "plan", "drafts", "manifest", "locations"):
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
                elif args.command == "audit":
                    result = audit(canon, plan, args.drafts, load(args.manifest), args.root)
                elif args.command == "index-artifacts":
                    result = index_artifacts(canon, plan, load(args.manifest), args.root)
                    save(args.out, result)
                elif args.command in ("verify-provenance", "coverage"):
                    manifest = load(args.manifest)
                    provenance = verify_provenance(canon, plan, args.drafts, manifest, args.root,
                                                   load(args.locations))
                    result = (provenance if args.command == "verify-provenance" else
                              coverage_report(canon, plan, args.drafts, provenance))
                    save(args.out, result)
                else:
                    manifest = load(args.manifest)
                    audit(canon, plan, args.drafts, manifest, args.root)
                    result = access_view(canon, plan, manifest,
                                         load(Path(args.root) / "build" / "provenance.json"),
                                         args.principal, args.as_of)
                    save(args.out, result)
        print(json.dumps(result, indent=2))
        return 0
    except (Invalid, OSError, ValueError, TypeError, KeyError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
