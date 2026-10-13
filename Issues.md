# Issues

| Severity | Finding | Impact |
|---|---|---|
| Medium | **Native PowerPoint changes chart meaning** — `.github/skills/company-content-generator/scripts/office_com_powerpoint_export_worker.ps1:238-265`; documented chart kinds are in `.github/skills/company-content-generator/references/contracts.md:305-313`. The native renderer ignores `kind` and draws every chart as bars. It also uses absolute values, making negative values appear as positive-direction bars. | Charts can communicate a different visualization or direction than the source data. Implement the documented chart semantics or reject unsupported chart inputs before export. |
