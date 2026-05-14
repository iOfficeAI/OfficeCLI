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
        var command = new Command("hwp", "Show experimental HWP/rhwp setup, recipes, sidecar notes, and coverage policy");
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
                "officecli create file.hwp --json",
                "officecli view file.hwp text --json",
                "officecli view file.hwp svg --page 1 --json",
                "officecli view file.hwp fields --json",
                "officecli view file.hwp field --field-name 회사명 --json",
                "officecli set file.hwp /field --prop name=회사명 --prop value=리지 --prop output=out.hwp --json",
                "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json",
                "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --in-place --backup --verify --json",
                "officecli set file.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=out.hwp --json",
                "officecli set file.hwpx /save-as-hwp --prop output=out.hwp --json"
            ]);
            var setup = HwpCapabilityJsonMapper.ToJsonArray([
                "run ./dev-install.sh to install officecli with rhwp sidecars",
                "optional: export OFFICECLI_HWP_ENGINE=rhwp-experimental",
                "export OFFICECLI_RHWP_BIN=/path/to/rhwp",
                "export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge.dll",
                "export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge"
            ]);
            var unsupported = HwpCapabilityJsonMapper.ToJsonArray([
                "silently falling back to HWPX when the requested native .hwp operation is ready",
                "ungated writes when the required rhwp sidecar or safe-save readback is unavailable",
                "claiming corpus-wide round-trip fidelity without fixture evidence"
            ]);
            var requiredEnv = new JsonArray();
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_HWP_ENGINE", "rhwp-experimental", "optional override; packaged sidecars are auto-discovered when present"));
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_RHWP_BIN", "/path/to/rhwp", "optional stock rhwp CLI fallback for text/SVG read-render"));
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_RHWP_BRIDGE_PATH", "/path/to/rhwp-officecli-bridge.dll", "optional explicit C# bridge path invoked by OfficeCLI"));
            requiredEnv.Add((JsonNode?)BuildEnvSpec("OFFICECLI_RHWP_API_BIN", "/path/to/rhwp-field-bridge", "optional explicit Rust rhwp API sidecar path for create/fields/text/table/export"));
            var diagnostics = HwpCapabilityJsonMapper.ToJsonArray([
                "officecli hwp doctor --json",
                "officecli capabilities --json",
                "officecli view file.hwp text --json"
            ]);
            var recipes = new JsonObject
            {
                ["createBlank"] = "officecli create file.hwp --json",
                ["readText"] = "officecli view file.hwp text --json",
                ["renderSvg"] = "officecli view file.hwp svg --page 1 --json",
                ["listFields"] = "officecli view file.hwp fields --json",
                ["readField"] = "officecli view file.hwp field --field-name 회사명 --json",
                ["setField"] = "officecli set file.hwp /field --prop name=회사명 --prop value=리지 --prop output=out.hwp --json",
                ["replaceText"] = "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json",
                ["replaceTextInPlace"] = "officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --in-place --backup --verify --json",
                ["setTableCell"] = "officecli set file.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=out.hwp --json",
                ["saveAsHwp"] = "officecli set file.hwpx /save-as-hwp --prop output=out.hwp --json"
            };
            var policies = HwpCapabilityJsonMapper.ToJsonArray([
                "default mutation mode writes to --prop output=<path>",
                "in-place text replacement is opt-in only and requires --in-place --backup --verify",
                "never combine --in-place with --prop output=<path>",
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
        writer.WriteLine("  ./dev-install.sh");
        writer.WriteLine("  # optional explicit overrides:");
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
        writer.WriteLine("  officecli create file.hwp --json");
        writer.WriteLine("  officecli view file.hwp text --json");
        writer.WriteLine("  officecli view file.hwp svg --page 1 --json");
        writer.WriteLine("  officecli view file.hwp fields --json");
        writer.WriteLine("  officecli view file.hwp field --field-name 회사명 --json");
        writer.WriteLine();
        writer.WriteLine("Mutation default writes a new output file:");
        writer.WriteLine("  officecli set file.hwp /field --prop name=회사명 --prop value=리지 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwp /field --prop id=1584999796 --prop value=리지 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwp /table/cell --prop section=0 --prop parent-para=3 --prop control=0 --prop cell=0 --prop value=오피스셀 --prop output=out.hwp --json");
        writer.WriteLine("  officecli set file.hwpx /save-as-hwp --prop output=out.hwp --json");
        writer.WriteLine();
        writer.WriteLine("Safe in-place text replacement (experimental, creates backup + manifest first):");
        writer.WriteLine("  officecli set file.hwp /text --prop find=마케팅 --prop value=브릿지 --in-place --backup --verify --json");
        writer.WriteLine();
        writer.WriteLine("Sidecar binaries used by the experimental bridge:");
        writer.WriteLine("  rhwp-officecli-bridge       C# bridge for OfficeCLI HWP routing");
        writer.WriteLine("  rhwp-field-bridge           Rust rhwp API sidecar for create/read/render/fields/text/table/export");
        writer.WriteLine();
        writer.WriteLine("Coverage policy:");
        writer.WriteLine("  - operation-gated rhwp coverage; if rhwp can do it and sidecars expose it, OfficeCLI should wire it");
        writer.WriteLine("  - experimental until fixture/corpus evidence promotes each operation");
        writer.WriteLine("  - default mutations write to --prop output=<path>");
        writer.WriteLine("  - in-place text replacement requires --in-place --backup --verify");
        writer.WriteLine("  - never combine --in-place with --prop output=<path>");
        writer.WriteLine("  - table cell mutation currently uses explicit rhwp coordinates; broader table discovery should be wired when rhwp exposes it");
        writer.WriteLine("  - verify outputs with text/SVG/Hancom round-trip evidence before relying on them");
        return true;
    }

    private static JsonObject BuildHwpDoctorReport()
    {
        var runtime = HwpRuntimeProbe.Probe();
        var checks = new JsonArray();
        checks.Add((JsonNode?)BuildRuntimeCheck(
            "OFFICECLI_HWP_ENGINE",
            runtime.EngineRequested || runtime.BridgeAvailable || runtime.ApiAvailable || runtime.RhwpAvailable,
            Environment.GetEnvironmentVariable("OFFICECLI_HWP_ENGINE"),
            "optional when rhwp sidecars are installed beside officecli",
            "export OFFICECLI_HWP_ENGINE=rhwp-experimental"));
        checks.Add((JsonNode?)BuildRuntimeCheck(
            "rhwp-officecli-bridge",
            runtime.BridgeAvailable,
            runtime.BridgePath,
            "C# bridge sidecar for existing-file read/render/mutation; not required for blank .hwp create",
            "run ./dev-install.sh or export OFFICECLI_RHWP_BRIDGE_PATH=/path/to/rhwp-officecli-bridge"));
        checks.Add((JsonNode?)BuildRuntimeCheck(
            "rhwp-field-bridge",
            runtime.ApiAvailable,
            runtime.ApiPath,
            "Rust rhwp API sidecar for .hwp create/mutate/export and direct read/render",
            "run ./dev-install.sh or export OFFICECLI_RHWP_API_BIN=/path/to/rhwp-field-bridge"));
        checks.Add((JsonNode?)BuildRuntimeCheck(
            "read-render runtime",
            runtime.ApiAvailable || runtime.RhwpAvailable,
            runtime.ApiPath ?? runtime.RhwpPath,
            "rhwp-field-bridge is preferred; stock rhwp is accepted only as read/render fallback",
            "install rhwp-field-bridge beside officecli, or export OFFICECLI_RHWP_BIN=/path/to/rhwp"));

        var operations = new JsonObject
        {
            ["read_text"] = BuildDoctorOperation(
                runtime.ReadRenderAvailable,
                "existing .hwp text extraction",
                runtime.ReadRenderAvailable ? null : "requires rhwp-officecli-bridge plus rhwp-field-bridge or stock rhwp",
                "officecli view file.hwp text --json"),
            ["render_svg"] = BuildDoctorOperation(
                runtime.ReadRenderAvailable,
                "existing .hwp SVG/page render",
                runtime.ReadRenderAvailable ? null : "requires rhwp-officecli-bridge plus rhwp-field-bridge or stock rhwp",
                "officecli view file.hwp svg --page 1 --json"),
            ["mutate_output"] = BuildDoctorOperation(
                runtime.MutationAvailable,
                "existing .hwp output-first field/text/table mutation",
                runtime.MutationAvailable ? null : "requires rhwp-officecli-bridge plus rhwp-field-bridge",
                "officecli set file.hwp /text --prop find=OLD --prop value=NEW --prop output=out.hwp --json"),
            ["create_blank"] = BuildDoctorOperation(
                runtime.CreateBlankAvailable,
                "new blank binary .hwp creation",
                runtime.CreateBlankAvailable ? null : "requires rhwp-field-bridge",
                "officecli create file.hwp --json")
        };

        var ok = runtime.ReadRenderAvailable || runtime.MutationAvailable || runtime.CreateBlankAvailable;
        return new JsonObject
        {
            ["schemaVersion"] = 1,
            ["topic"] = "hwp-rhwp-doctor",
            ["ok"] = ok,
            ["autoDiscovery"] = new JsonObject
            {
                ["bridgePath"] = runtime.BridgePath,
                ["apiPath"] = runtime.ApiPath,
                ["rhwpPath"] = runtime.RhwpPath,
                ["readRenderAvailable"] = runtime.ReadRenderAvailable,
                ["mutationAvailable"] = runtime.MutationAvailable,
                ["createBlankAvailable"] = runtime.CreateBlankAvailable
            },
            ["operations"] = operations,
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

    private static JsonObject BuildRuntimeCheck(
        string name,
        bool ok,
        string? value,
        string detail,
        string hint)
        => new()
        {
            ["name"] = name,
            ["ok"] = ok,
            ["isSet"] = !string.IsNullOrWhiteSpace(value),
            ["value"] = value,
            ["detail"] = detail,
            ["hint"] = ok ? null : hint
        };

    private static JsonObject BuildDoctorOperation(
        bool ready,
        string detail,
        string? blockedBy,
        string example)
        => new()
        {
            ["ready"] = ready,
            ["detail"] = detail,
            ["blockedBy"] = blockedBy,
            ["example"] = example
        };

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
        foreach (var op in report["operations"]!.AsObject())
        {
            var opData = op.Value!.AsObject();
            var mark = opData["ready"]?.GetValue<bool>() == true ? "ready" : "blocked";
            writer.WriteLine($"  {mark}: {op.Key}");
            var blockedBy = opData["blockedBy"]?.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(blockedBy))
                writer.WriteLine($"       {blockedBy}");
        }
        writer.WriteLine();
        writer.WriteLine($"Next: {report["nextCommand"]?.GetValue<string>()}");
    }
}
