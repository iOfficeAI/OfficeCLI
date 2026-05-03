// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli;

static partial class CommandBuilder
{
    private static string BuildViewDescription()
        => """
            View document in different modes.

            Agent examples:
              officecli view report.docx text --json
              officecli view sheet.xlsx text --json
              officecli view file.hwp text --json
              officecli view file.hwp svg --page 1 --json
              officecli view file.hwp fields --json
              officecli view file.hwp field --field-name 회사명 --json

            HWP requires OFFICECLI_HWP_ENGINE=rhwp-experimental plus bridge paths; run `officecli help hwp`.
            """;
}
