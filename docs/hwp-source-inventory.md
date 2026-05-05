# HWP/HWPX Source Inventory

This inventory supports the HWP/HWPX capability contract. Mutable web sources require
a reproducibility identifier. Sources without one can support background context only,
not OfficeCLI capability claims.

| Source | URL | Accessed | Observed version or commit | Retrieved artifact hash or note | Claim allowed in OfficeCLI docs |
|---|---|---|---|---|---|
| Hancom Tech HWPX format structure | https://tech.hancom.com/hwpxformat/ | 2026-05-03 KST | web page | no local archive in Phase 0; background context only | HWPX is ZIP/XML-based; this does not prove Hancom-compatible writing. |
| Hancom HWPX FAQ | https://www.hancom.com/support/faqCenter/faq/detail/2784 | 2026-05-03 KST | web page | no local archive in Phase 0; background context only | HWPX is an OWPML-based open document format registered as KS X 6101. |
| rhwp repository | https://github.com/edwardkim/rhwp | 2026-05-03 KST | `0fb3e6758b8ad11d2f3c3849c83b914684e83863` | `git ls-remote ... HEAD` | Candidate upstream HWP/HWPX read/render/edit engine only; no OfficeCLI support claim. |
| HOP repository | https://github.com/golbin/hop | 2026-05-03 KST | `bd6839bf55f8c2819a61c120421be60c4074e2a3` | `git ls-remote ... HEAD` | Candidate desktop integration evidence only; no OfficeCLI support claim. |
| HOP development note | https://github.com/golbin/hop/blob/main/docs/DEVELOPMENT.md | 2026-05-03 KST | `bd6839bf55f8c2819a61c120421be60c4074e2a3` | `git ls-remote ... HEAD`; exact file content not archived in Phase 0 | Upstream HWPX save limitations remain evidence against blanket write claims. |
| Microsoft .NET single-file deployment | https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file/overview | 2026-05-03 KST | web page | no local archive in Phase 0; packaging context only | Single-file publish is RID-specific and native-library behavior must be tested. |
| Wasmtime .NET embedding | https://bytecodealliance.github.io/wasmtime-dotnet/articles/intro.html | 2026-05-03 KST | web page | no local archive in Phase 0; packaging context only | Wasmtime.NET is an embedding option only; packaging and latency must be measured. |

Rule: upstream rhwp/HOP claims are candidate capability evidence only. They are not
OfficeCLI support claims until `officecli capabilities --json` returns
`roundtrip-verified` with fixture evidence and Hancom-compatible evidence where
required.
