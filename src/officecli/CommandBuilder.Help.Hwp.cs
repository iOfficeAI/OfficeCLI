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
        command.Add(BuildHwpDoctorCommand(jsonOption));

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

    private static Command BuildHwpDoctorCommand(Option<bool> jsonOption)
    {
        var command = new Command("doctor", "Check experimental HWP/rhwp environment readiness");
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            return SafeRun(() =>
            {
                var report = BuildHwpDoctorReport();
                if (json)
                {
                    writerJson(Console.Out, report);
                    return report["ok"]?.GetValue<bool>() == true ? 0 : 2;
                }

                WriteHwpDoctorText(Console.Out, report);
                return report["ok"]?.GetValue<bool>() == true ? 0 : 2;
            }, json);
        });

        return command;

        static void writerJson(TextWriter writer, JsonObject report)
        {
            writer.WriteLine(OutputFormatter.WrapEnvelope(report.ToJsonString(OutputFormatter.PublicJsonOptions)));
        }
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
            var requiredEnv = new JsonArray();
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_HWP_ENGINE", "rhwp-experimental", "selects the experimental rhwp bridge"));
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_RHWP_BIN", "/path/to/rhwp", "stock rhwp CLI for text/SVG read-render"));
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_RHWP_BRIDGE_PATH", "/path/to/rhwp-officecli-bridge.dll", "C# bridge invoked by OfficeCLI"));
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_RHWP_API_BIN", "/path/to/rhwp-field-bridge", "Rust rhwp API sidecar for fields/text/table mutation"));
            var diagnostics = HwpCapabilityJsonMapper.ToJsonArray([
                "officecli hwp doctor --json",
                "officecli capabilities --json",
                "officecli view file.hwp text --json"
            ]);
            var recipes = new JsonObject
            {
                ["readText"] = "officecli view file.hwp text --json",
                ["renderSvg"] = "officecli view file.hwp svg --page 1 --json",
                ["listFields"] = "officecli view file.hwp fields --json",
                ["readField"] = "officecli view file.hwp field --field-name 회사명 --json",
                ["setField"] = "officecli set file.hwp /field --prop name=회사명 --prop value=리지 --prop output=out.hwp --json",
                ["replaceText"] = "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json",
                ["setTableCell"] = "officecli set file.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=out.hwp --json"
            };
            var policies = HwpCapabilityJsonMapper.ToJsonArray([
                "never overwrite binary .hwp in place",
                "always write mutations to --prop output=<path>",
                "run officecli hwp doctor --json before HWP work",
                "verify outputs with text/SVG/Hancom evidence before relying on them"
            ]);
            var data = new JsonObject
            {
                ["schemaVersion"] = 1,
                ["topic"] = "hwp-rhwp",
                ["status"] = "experimental",
                ["setup"] = setup,
                ["requiredEnv"] = requiredEnv,
                ["setupCommands"] = setup.DeepClone(),
                ["commands"] = commands,
                ["recipes"] = recipes,
                ["diagnostics"] = diagnostics,
                ["policies"] = policies,
                ["unsupported"] = unsupported,
                ["capabilityProbe"] = "officecli capabilities --json",
                ["doctor"] = "officecli hwp doctor --json"
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
        writer.WriteLine("  officecli hwp doctor --json");
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
        writer.WriteLine("  - always write mutations to --prop output=<path>");
        writer.WriteLine("  - table cell mutation needs explicit rhwp coordinates");
        writer.WriteLine("  - verify outputs with text/SVG/Hancom round-trip evidence before relying on them");
        return true;
    }

    private static JsonObject BuildHwpDoctorReport()
    {
        var checks = new JsonArray();
        checks.Add((JsonNode?)BuildEnvCheck(
                "OFFICECLI_HWP_ENGINE",
                "rhwp-experimental",
                requireFile: false,
                "export OFFICECLI_HWP_ENGINE=rhwp-experimental"));
        checks.Add((JsonNode?)BuildEnvCheck(
                "OFFICECLI_RHWP_BIN",
                expectedValue: null,
                requireFile: true,
                "export OFFICECLI_RHWP_BIN=/path/to/rhwp"));
        checks.Add((JsonNode?)BuildEnvCheck(
                "OFFICECLI_RHWP_BRIDGE_PATH",
                expectedValue: null,
                requireFile: true,
                "export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge.dll"));
        checks.Add((JsonNode?)BuildEnvCheck(
                "OFFICECLI_RHWP_API_BIN",
                expectedValue: null,
                requireFile: true,
                "export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge"));

        var ok = checks.OfType<JsonObject>().All(check => check["ok"]?.GetValue<bool>() == true);
        return new JsonObject
        {
            ["schemaVersion"] = 1,
            ["topic"] = "hwp-rhwp-doctor",
            ["ok"] = ok,
            ["checks"] = checks,
            ["nextCommand"] = ok ? "officecli capabilities --json" : "officecli help hwp",
            ["capabilityProbe"] = "officecli capabilities --json"
        };
    }

    private static JsonObject BuildEnvSpec(string name, string example, string purpose)
        => new()
        {
            ["name"] = name,
            ["example"] = example,
            ["purpose"] = purpose
        };

    private static JsonObject BuildEnvCheck(string name, string? expectedValue, bool requireFile, string hint)
    {
        var value = Environment.GetEnvironmentVariable(name);
        var isSet = !string.IsNullOrWhiteSpace(value);
        var expectedMatches = expectedValue is null
            || string.Equals(value, expectedValue, StringComparison.OrdinalIgnoreCase);
        var fileExists = !requireFile || (isSet && File.Exists(value));
        var ok = isSet && expectedMatches && fileExists;

        return new JsonObject
        {
            ["name"] = name,
            ["ok"] = ok,
            ["isSet"] = isSet,
            ["value"] = value,
            ["expected"] = expectedValue,
            ["requiresExistingFile"] = requireFile,
            ["fileExists"] = requireFile ? fileExists : null,
            ["hint"] = ok ? null : hint
        };
    }

    private static void WriteHwpDoctorText(TextWriter writer, JsonObject report)
    {
        var ok = report["ok"]?.GetValue<bool>() == true;
        writer.WriteLine(ok ? "HWP/rhwp doctor: OK" : "HWP/rhwp doctor: NOT READY");
        writer.WriteLine();
        foreach (var check in report["checks"]!.AsArray().OfType<JsonObject>())
        {
            var mark = check["ok"]?.GetValue<bool>() == true ? "ok" : "missing";
            writer.WriteLine($"  {mark}: {check["name"]?.GetValue<string>()}");
            var hint = check["hint"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(hint))
                writer.WriteLine($"       {hint}");
        }
        writer.WriteLine();
        writer.WriteLine($"Next: {report["nextCommand"]?.GetValue<string>()}");
    }
}
