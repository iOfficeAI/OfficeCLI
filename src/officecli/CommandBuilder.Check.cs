// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildValidateCommand(Option<bool> jsonOption)
    {
        var validateFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var validateCommand = new Command("validate", "Validate document against OpenXML schema");
        validateCommand.Add(validateFileArg);
        validateCommand.Add(jsonOption);
        validateCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(validateFileArg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "validate";
                req.Json = json;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var errors = handler.Validate();
            if (json)
            {
                var validationJson = FormatValidationErrors(errors);
                Console.WriteLine(OutputFormatter.WrapEnvelope(validationJson));
            }
            else
            {
                if (errors.Count == 0)
                {
                    Console.WriteLine("Validation passed: no errors found.");
                }
                else
                {
                    // R7-bt-4: schema validation reports go to stderr —
                    // callers piping `validate` for CI gates need to see
                    // the failure summary on the diagnostic stream rather
                    // than mixed into stdout. Mirrors the resident path.
                    Console.Error.WriteLine($"Found {errors.Count} validation error(s):");
                    foreach (var err in errors)
                    {
                        Console.Error.WriteLine($"  [{err.ErrorType}] {err.Description}");
                        if (err.Path != null) Console.Error.WriteLine($"    Path: {err.Path}");
                        if (err.Part != null) Console.Error.WriteLine($"    Part: {err.Part}");
                    }
                }
            }
            return errors.Count > 0 ? 1 : 0;
        }, json); });

        return validateCommand;
    }
}
