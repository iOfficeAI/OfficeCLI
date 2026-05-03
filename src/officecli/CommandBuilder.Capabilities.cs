// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildCapabilitiesCommand(Option<bool> jsonOption)
    {
        var command = new Command("capabilities", "Show machine-readable OfficeCLI capability information");
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            return SafeRun(() =>
            {
                var report = HwpCapabilityFactory.BuildReport();
                if (json)
                {
                    Console.WriteLine(HwpCapabilityJsonMapper.BuildEnvelope(report).ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
                    return 0;
                }

                Console.WriteLine($"OfficeCLI capabilities schema {report.SchemaVersion}");
                Console.WriteLine("HWP/HWPX support is gated by `officecli capabilities --json`.");
                Console.WriteLine("Run `officecli help hwp` for rhwp bridge setup, examples, and support boundaries.");
                Console.WriteLine("Binary .hwp in-place overwrite/save-as is unsupported; experimental mutations require output=<path>.");
                return 0;
            }, json);
        });

        return command;
    }
}
