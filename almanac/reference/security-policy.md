---
title: "Security Policy"
summary: "Lookup reference for OfficeCLI vulnerability reporting and the repository's file-opening, resource-limit, and remote-fetch guards."
topics: [reference, security]
sources:
  - id: security-policy
    type: file
    path: SECURITY.md
  - id: handler-factory
    type: file
    path: src/officecli/Handlers/DocumentHandlerFactory.cs
  - id: ssrf-guard
    type: file
    path: src/officecli/Core/SsrfGuard.cs
  - id: document-limits
    type: file
    path: src/officecli/Core/DocumentLimits.cs
---

OfficeCLI's security policy combines private vulnerability reporting with runtime guards for untrusted Office files and caller-supplied remote URLs. The public policy asks reporters to use GitHub private vulnerability reporting, while the code centralizes file admission, decompression-bomb checks, recursion limits, regex timeouts, and SSRF protection in shared helpers [@security-policy] [@handler-factory] [@document-limits] [@ssrf-guard]. The opening path is part of the [document handler lifecycle](../architecture/handlers/document-handler-lifecycle), and plugin resolution is covered by the [plugin system](../architecture/plugins/plugin-system).

## Reporting

Security vulnerabilities should be reported privately through GitHub's "Security -> Report a vulnerability" flow, not as public issues [@security-policy]. Reports should include a description and impact, reproduction steps or a minimal sample file, the OfficeCLI version from `officecli --version`, and the reporter's operating system [@security-policy].

Security fixes are applied to the latest released version, and reporters are asked to upgrade to the latest version before reporting [@security-policy].

## File admission

`DocumentHandlerFactory.Open` is the shared gate for document opening [@handler-factory]. It rejects empty paths with code `file_required`, missing paths with code `file_not_found`, and zero-byte files with code `corrupt_file` before handler construction [@handler-factory].

For native `.docx`, `.xlsx`, and `.pptx` files, the factory checks the zip directory before the Open XML SDK opens the package [@handler-factory]. It rejects archives with more than `100000` entries, more than `2 GiB` total uncompressed data, or an overall compression ratio above `1000x` once compressed data exceeds 64 KiB [@handler-factory] [@document-limits].

## Repair before open

The factory repairs two producer defects centrally. It strips dangling internal package relationships before open when native packages contain relationships to missing parts, and it rewrites unsupported XML encoding declarations to UTF-8 before retrying [@handler-factory]. These repairs are applied before dispatch to native or plugin-backed handlers [@handler-factory].

## Resource limits

`DocumentLimits` is the shared source for denial-of-service limits [@document-limits].

| Limit | Value | Guarded risk |
|---|---:|---|
| `MaxRecursionDepth` | `256` | Deep document trees or formulas exhausting process stack [@document-limits] |
| `MaxUncompressedBytes` | `2 GiB` | OOXML zip decompression bombs [@document-limits] |
| `MaxZipEntries` | `100000` | Archives with excessive entry counts [@document-limits] |
| `MaxCompressionRatio` | `1000` | Highly compressed crafted archives [@document-limits] |
| `RegexMatchTimeout` | `5 seconds` | Catastrophic backtracking in user-supplied regex patterns [@document-limits] |

`EnsureDepth` throws a `CliException` with code `max_depth_exceeded` when depth exceeds the limit or the runtime stack probe reports insufficient remaining execution stack [@document-limits].

## Remote fetch guard

`SsrfGuard` is the shared HTTP/HTTPS protection for remote image and file sources [@ssrf-guard]. Its guarded handler validates the actual IP address in `ConnectCallback` for each connection, including redirect hops, so DNS rebinding between pre-resolution and connect is not accepted [@ssrf-guard].

Only globally routable addresses are allowed. The guard blocks loopback, unspecified, private IPv4 ranges, link-local ranges including `169.254.0.0/16`, CGNAT, multicast or reserved IPv4, IPv6 link-local, site-local, multicast, and IPv6 unique-local `fc00::/7` addresses [@ssrf-guard].

Remote response bodies are bounded by `SsrfGuard.MaxRemoteBytes`, which is `100 MB` [@ssrf-guard]. `ReadBounded` refuses reads that exceed that limit even when the server omits or lies about `Content-Length` [@ssrf-guard].
