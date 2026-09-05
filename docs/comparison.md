# How we compare

We are a **new project**, so we don't compete on star count - we compete on the gap nobody fills: a **single, local-first, portable, enterprise-grade document MCP** that covers the full lifecycle. Here's the field, with real numbers pulled from the GitHub API (Sep 2026).

## Competitive landscape

| Project | Stars | Language | Platform | Last push | Focus & honest limit |
|---|---|---|---|---|---|
| [doc-ops-mcp](https://github.com/tele-ai/doc-ops-mcp) (Tele-AI) | 139 | TypeScript | Multi | 2026-08-31 | Conversion (PDF/DOCX/HTML/MD) + watermark. No security/compliance layer, no books/pivots/guidance. |
| [opendocwork-mcp](https://github.com/Aimino-Tech/opendocwork-mcp) (Aimino) | ~154* | Rust | Multi, local-first | 2026-05-26 | "Open-source Aspose": Excel/Word/PPT/PDF + pivot/charts. Strong rival; less enterprise-security surface, Rust harder to extend for the Hermes community. |
| [OfficeMCP](https://github.com/OfficeMCP/OfficeMCP) | 116 | Python/.NET COM | **Windows only** | 2025-05-28 stale ~1y | Drives real MS Office via COM + RunPython (arbitrary code execution). Doesn't run on Linux/Mac; not maintained. |
| [mcp-ms-office-documents](https://github.com/ForLegalAI/mcp-ms-office-documents) (ForLegalAI) | 39 | Python/Docker | Docker+HTTP | 2026-09-05 active | One-shot generation from Markdown (PPT/Word/Excel/email). Not a full read/edit/security lifecycle. |
| mcp-docgen (Touka) | new (PyPI v0.8) | Python | Multi, local | active | Markdown->Office read/edit/convert with honest fidelity. Closest in design philosophy; lacks security/compliance + guidance + books. |
| office-mcp / docforge-mcp / mhackermsft OfficeMCP | 0-4 each | Python/C# | varies | recent/niche | Granular CRUD (47 tools) or forks of ForLegalAI. Narrower scope, higher context cost. |

\* opendocwork star count from search index; exact figure rate-limited at capture time - marked "~" rather than asserted.

## Why The Office Worker is the one to install

1. **The only full-lifecycle enterprise + local + portable package.** Create -> read -> extract structured (MD/JSON) -> convert <-> -> secure (redact / scrub_metadata / protect_office / flatten) -> batch -> sign (real PAdES) -> audit (document_diff / verify_pdf_signature) -> pivot/charts/books - in one multiplatform local package. No single competitor spans all of it.
2. **Truly cross-platform vs OfficeMCP.** The #3 by stars runs Windows-only and is ~1 year stale. We ship green CI on Linux + Mac, proven on ARM64.
3. **Security-first vs arbitrary-code rivals.** OfficeMCP exposes `RunPython(codes)` = arbitrary code execution. We use deterministic `office_batch` + `safe_out`, no free code path - a strong argument for enterprise adoption.
4. **Measured context efficiency.** ~2550 tok/turn across 29 specialized tools vs ~6400 for a 47-tool CRUD suite. Less bloat = more stable agents and lower cost. This is measured, not claimed.
5. **Honest fidelity + proactive guidance.** Only mcp-docgen reports honest fidelity; we add `next_steps` (orients the agent to the next logical action) and `owi doctor` (assisted onboarding). That puts us ahead even of the closest design rival in agent UX.
6. **Zero API keys / privacy-local as a flag.** Data sovereignty is explicit, not an afterthought - documents never leave your machine.

> Honest positioning: we win on the *gap*, not on existing adoption. Distribution (MCP directories, Show HN / Reddit, README demos) is what converts that gap into installs - not more features.

