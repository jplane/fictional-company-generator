import json
import os
import sys
from collections import Counter
from pathlib import Path


def main():
    report = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    completeness = report["analysis_completeness"]
    coverage = completeness["coverage_percent"]
    complete = completeness["is_complete"]
    summary = [
        "## SkillSpector static scan",
        "",
        f"- Unsuppressed risk score: {report['risk_assessment']['score']}/100 (blocks above 50).",
        f"- Reviewed findings suppressed: {report['suppressed_count']}.",
        f"- Coverage: {coverage}% ({'complete' if complete else 'INCOMPLETE; nonblocking'}).",
        "",
    ]
    reasons = Counter(item["reason_code"] for item in completeness["ledger_exceptions"])
    summary.extend(f"- `{reason}`: {count}" for reason, count in sorted(reasons.items()))
    summary.extend([
        "",
        "See the skillspector-report artifact for findings, suppressed evidence, "
        "coverage limitations, and revisions. A passing scan is not security certification.",
        "",
    ])
    with Path(os.environ["GITHUB_STEP_SUMMARY"]).open("a", encoding="utf-8") as stream:
        stream.write("\n".join(summary))
    if report["execution_successful"] is not True or completeness["execution_successful"] is not True:
        raise SystemExit("SkillSpector execution failed; inspect the report artifact.")
    if not complete:
        print(f"::warning title=SkillSpector incomplete coverage::Static analysis coverage is "
              f"{coverage}%. Coverage is nonblocking; inspect the job summary and report artifact.")


if __name__ == "__main__":
    main()
