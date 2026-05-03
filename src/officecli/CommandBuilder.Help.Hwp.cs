// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;
using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildHwpHelpCommand(Option<bool> jsonOption)
    {
        var command = new Command("hwp", "Show experimental HWP/rhwp setup, recipes, sidecar notes, and boundaries");
        command.Aliases.Add("rhwp");
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            return SafeRun(() =>
            {
                WriteHwpBridgeHelp("hwp", Console.Out, json);
                return 0;
            }, json);
        });

        return command;
    }

    private static bool WriteHwpBridgeHelp(string topic, TextWriter writer, bool json)
    {
        if (!string.Equals(topic, "hwp", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(topic, "rhwp", StringComparison.OrdinalIgnoreCase))
            return false;

        if (json)
        {
            var commands = HwpCapabilityJsonMapper.ToJsonArray([
                "officecli capabilities --json",
                "officecli view file.hwp text --json",
                "officecli view file.hwp svg --page 1 --json",
                "officecli view file.hwp fields --json",
                "officecli view file.hwp field --field-name 회사명 --json",
                "officecli set file.hwp /field --prop name=회사명 --prop value=리지 --prop output=out.hwp --json",
                "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json",
                "officecli set file.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=out.hwp --json"
            ]);
            var setup = HwpCapabilityJsonMapper.ToJsonArray([
                "export OFFICECLI_HWP_ENGINE=rhwp-experimental",
                "export OFFICECLI_RHWP_BIN=/path/to/rhwp",
                "export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge.dll",
                "export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge"
            ]);
            var unsupported = HwpCapabilityJsonMapper.ToJsonArray([
                "in-place binary .hwp overwrite",
                "production default HWP engine",
                "arbitrary table discovery/mutation without explicit coordinates",
                "claiming corpus-wide round-trip fidelity without fixtures"
            ]);
            var data = new JsonObject
            {
                ["topic"] = "hwp-rhwp",
                ["status"] = "experimental",
                ["setup"] = setup,
                ["commands"] = commands,
                ["unsupported"] = unsupported,
                ["capabilityProbe"] = "officecli capabilities --json"
            };
            writer.WriteLine(OutputFormatter.WrapEnvelope(data.ToJsonString(OutputFormatter.PublicJsonOptions)));
            return true;
        }

        writer.WriteLine("HWP / rhwp Bridge Help (experimental)");
        writer.WriteLine();
        writer.WriteLine("Setup:");
        writer.WriteLine("  export OFFICECLI_HWP_ENGINE=rhwp-experimental");
        writer.WriteLine("  export OFFICECLI_RHWP_BIN=/path/to/rhwp");
        writer.WriteLine("  export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge.dll");
        writer.WriteLine("  export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge");
        writer.WriteLine();
        writer.WriteLine("Probe:");
        writer.WriteLine("  officecli capabilities --json");
        writer.WriteLine();
        writer.WriteLine("Top-level help:");
        writer.WriteLine("  officecli hwp");
        writer.WriteLine("  officecli rhwp");
        writer.WriteLine();
        writer.WriteLine("Read/render:");
        writer.WriteLine("  officecli view file.hwp text --json");
        writer.WriteLine("  officecli view file.hwp svg --page 1 --json");
        writer.WriteLine("  officecli view file.hwp fields --json");
        writer.WriteLine("  officecli view file.hwp field --field-name 회사명 --json");
        writer.WriteLine();
        writer.WriteLine("Mutation writes a new output file; it does not overwrite the input:");
        writer.WriteLine("  officecli set file.hwp /field --prop name=회사명 --prop value=리지 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwp /field --prop id=1584999796 --prop value=리지 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=out.hwp --json");
        writer.WriteLine();
        writer.WriteLine("Sidecar binaries used by the experimental bridge:");
        writer.WriteLine("  OFFICECLI_RHWP_BRIDGE_PATH  C# bridge for rhwp CLI read/render");
        writer.WriteLine("  OFFICECLI_RHWP_API_BIN      Rust rhwp API sidecar for fields/text/table mutation");
        writer.WriteLine();
        writer.WriteLine("Boundaries:");
        writer.WriteLine("  - experimental only; do not claim production default HWP support");
        writer.WriteLine("  - binary .hwp in-place overwrite is unsupported");
        writer.WriteLine("  - table cell mutation needs explicit rhwp coordinates");
        writer.WriteLine("  - verify outputs with text/SVG/Hancom round-trip evidence before relying on them");
        return true;
    }
}
