# Layering a customer use case onto a general-purpose corpus

The core skill creates coherent enterprise content. It does not require a
decision question, benchmark, evaluator, or downstream agent. Its statements,
viewpoints, source history, optional corpus conditions, simulated visibility,
and portable provenance can support search, summarization, onboarding, knowledge
management, training, document workflows, analytics, and evaluation.

This guide describes a **separate application layer**. The snippets are
illustrative Markdown examples, not implemented schemas or generated customer
documents. Keep the wrapper and evaluator outside the core corpus root;
the core layout remains `canon.json`, `documents\`, and `build\`.
Never give independent answers or evaluation partitions to corpus writers.

## Fictional customer example

Cedarline Robotics is considering an acquisition of Harbor Automation.
Its acquisition team wants an assistant to assemble a dated evidence brief,
identify conflicting integration estimates, distinguish approved plans from
individual assumptions, and recommend whether to proceed to the next review.
This is an example, not a requirement to orient the skill around acquisitions.

| Use-case need | General-purpose core capability | Application-layer responsibility |
|---|---|---|
| Deal evidence | Canonical business records; custom genres; exact claim receipts | Select acquisition-specific subjects, questions, and input documents |
| Finance, engineering, operations perspectives | Canonical viewpoints and attributed statements | Choose relevant customer functions and decision authority |
| Estimates changing over time | Immutable source versions; corrections/retractions; effective/known dates | Define the brief's cutoff and interpretation of relevant history |
| Missing integration analysis | Explicit unknown/undocumented/omitted-source conditions | Decide what the absence means for this question |
| Differing integration costs | Explicit conflict, units/period scope, source attribution | Define escalation or reconciliation policy |
| A superseded planning deck | Explicit stale condition with reference date | Define when the use case should prefer a newer approved source |
| Team/public evidence boundaries | Author knowledge and simulated recipient views | Bind actual service principals and enforce deployment ACLs |
| Evidence references | Final-artifact locators, hashes, optional local/SharePoint bindings | Make citations navigable and verify deployed item bytes |
| Performance measurement | Delivered corpus coverage only | Independent truth, partitions, expected outcomes, scoring |

Translate requirements into supported request fields and canonical records
before freeze. With a fixed 40-document budget, required viewpoints/sources
occupy those 40 slots; no wrapper should silently inflate the core target.
Keep depth (`light`/`medium`/`deep`) separate from conditions and evaluation
difficulty. A deep clean acquisition corpus and a light conflicting corpus
are both valid.

## 1. External `DecisionScenario` object

The wrapper may define a question and workflow context:

```json
{
  "scenario_id": "cedarline-harbor-integration-review",
  "customer": "Cedarline Robotics",
  "decision_question": "Is the integration evidence sufficient for the next review?",
  "evidence_cutoff": "2025-06-30",
  "reader_profile": "acquisition-review-team",
  "corpus_id": "cedarline-harbor",
  "evidence_document_ids": ["doc-0007", "doc-0012"],
  "required_topics": ["migration_cost", "support_capacity"],
  "workflow": "integration-review"
}
```

This object belongs to the consuming application, **not** the core canon's
record kinds or request allow-list. The wrapper resolves the profile to actual
principals and core viewpoints/groups, selects an audited recipient view, and
verifies cutoff eligibility. Scenario-specific rules may define which evidence
is sufficient; the generator does not decide that.

The same corpus can support a different question, an onboarding summary, or
search without regenerating every document. A scenario need not expose the
full private canon or build receipts to its agent.

## 2. Independent evaluation truth

If evaluating an assistant, maintain a separate evaluator-controlled answer
set, including relevant facts, source conflicts, unresolved information, and
allowed conclusions:

```json
{
  "scenario_id": "cedarline-harbor-integration-review",
  "truth_revision": "reviewed-1",
  "accepted_findings": [
    "The cost estimates use different migration assumptions.",
    "Support capacity has not been established by the cutoff."
  ],
  "required_evidence": [
    {"document_id": "doc-0007", "claim_id": "migration-estimate"}
  ],
  "unresolved_topics": ["support_capacity"],
  "review_method": "Independent domain review against the frozen artifact set"
}
```

Canon truth and source assertions are not automatically a defensible answer
key. An objective business actual, a source's assertion, and an evaluator's
judgment about decision sufficiency are different things. An evaluator can
use provenance as an audit aid, but merely copying canon fields or reusing the
same generated rationale is not independence.

Keep truth inaccessible to writers and the evaluated agent. Review independence,
ambiguity, evidence cutoff, and acceptable alternative answers separately.
The core QA asks whether artifacts faithfully carry their modeled content,
not whether a customer decision is correct.

## 3. Development, held-out, and challenge sets

Partition **scenario families** in an external dataset registry:

```json
{
  "dataset_id": "integration-review-suite",
  "split_unit": "scenario-family",
  "assignments": [
    {"family_id": "harbor-integration", "split": "development"},
    {"family_id": "ridgeway-platform", "split": "held-out"},
    {"family_id": "northport-service", "split": "challenge"}
  ]
}
```

Do not randomly split documents from the same deal across development and
held-out sets and call the result independent. Shared names, figures, source
versions, and correction histories can leak answers. The wrapper owns family
identity, deduplication, access restrictions, registry versioning, and exposure
logging. Corpus IDs/fingerprints support traceability but do not prevent leakage.

Challenge membership is not `content_depth: deep`. Explicit conditions,
time cutoffs, viewpoint diversity, and workflow requirements can influence
difficulty; depth controls substantive richness. Keep one partition label out
of ordinary business prose and recipient manifests.

## 4. Expected dispositions and decision metrics

A customer-specific evaluator may define outcomes such as:

| Disposition | Illustrative Cedarline rule |
|---|---|
| Proceed to next review | Required evidence is current and sufficient under the customer's reviewed policy |
| Request clarification | A material scope/assumption conflict remains unresolved |
| Defer | A required capacity assessment is unavailable |
| Escalate | The issue exceeds the assistant's authority |

These labels and their thresholds are **not core defaults**. Other customers
may permit different outcomes, use a human-only recommendation workflow, or
have no decision task at all. An expected disposition is an external evaluation
target, not a requirement that the business documents announce that decision.
An actual historical company decision can still be a canonical `decision`
statement without carrying an expected agent disposition.

Possible external metrics include finding accuracy, evidence-location
accuracy, contradiction recognition, uncertainty calibration, abstention/
escalation quality, unsupported-claim rate, and disposition agreement.
Define denominators, cutoff handling, alternative correct answers, and reviewer
adjudication. Coverage counts and token matches are not decision-performance
scores; a source citation is not proof of sound reasoning.

## Deployment boundary

The wrapper binds logical identities to local files or uploaded drive/item IDs,
checks remote content against artifact hashes, and assigns actual ACLs.
It supplies only selected, permitted evidence to the agent and stores scenario
configuration, truth, split registries, and scoring outside that evidence set.
Core `access-view` is a simulation/selection tool, not real security enforcement
or an uploader.

For a non-evaluation use case, omit truth, partitions, dispositions, and metrics
entirely. For example, an onboarding application can select current
role-appropriate policies and product material, use source status to distinguish
draft guidance, and attach exact document citations. All four excluded
evaluation concepts remain unnecessary.
