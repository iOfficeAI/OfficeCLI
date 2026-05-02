# OfficeCLI Fork — Apache 2.0 Compliance Notes

## License Status
- Upstream: Apache License 2.0 (LICENSE file intact)
- Upstream NOTICE file: **none shipped** (checked v1.0.28)
- Apache 2.0 §4(b): modified files must carry prominent change notices

## Our Fork
- Fork owner: lidge-jun
- Fork URL: https://github.com/lidge-jun/OfficeCLI
- Upstream: https://github.com/iOfficeAI/OfficeCLI

## Planned Modifications
- CJK font handling (CjkHelper.cs) — Korean/Japanese/Chinese font metadata
- CJK language tag injection (w:lang, a:lang attributes)
- Kinsoku line-break processing
- East Asian character spacing

## Attribution (for distribution)
If we ship a bundled binary, include this NOTICE:

```
OfficeCLI
Copyright (c) iOfficeAI contributors

This product includes software developed at
iOfficeAI (https://github.com/iOfficeAI/OfficeCLI).

Licensed under the Apache License, Version 2.0.

Modifications by cli-jaw contributors:
- CJK (Korean/Japanese/Chinese) font handling and language tags
- CJK line-break (kinsoku) processing
- East Asian character spacing support
```

## Build Environment
- .NET 10.0.201 SDK installed at ~/.dotnet/
- `dotnet build src/officecli/officecli.csproj -c Release` — Build succeeded (0 warnings, 0 errors)
