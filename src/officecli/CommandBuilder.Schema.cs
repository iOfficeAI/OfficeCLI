// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Help;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildSchemaCommand(Option<bool> jsonOption)
    {
        var command = new Command("schema", "List and validate embedded OfficeCLI help schemas");
        command.Add(BuildSchemaListCommand(jsonOption));
        command.Add(BuildSchemaValidateCommand(jsonOption));
        return command;
    }

    private static Command BuildSchemaListCommand(Option<bool> jsonOption)
    {
        var formatOption = new Option<string?>("--format")
        {
            Description = "Restrict to one format: docx, xlsx, pptx, hwpx, or hwp",
        };
        var command = new Command("list", "List embedded schema-help entries");
        command.Add(formatOption);
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            return SafeRun(() =>
            {
                var format = result.GetValue(formatOption);
                var formats = string.IsNullOrWhiteSpace(format)
                    ? SchemaHelpLoader.ListFormats()
                    : new[] { SchemaHelpLoader.NormalizeFormat(format!) };
                var entries = BuildSchemaListEntries(formats);

                if (json)
                {
                    var data = new JsonObject
                    {
                        ["schemaVersion"] = 1,
                        ["formats"] = BuildFormatArray(formats),
                        ["entries"] = entries,
                    };
                    Console.WriteLine(OutputFormatter.WrapEnvelope(data.ToJsonString(OutputFormatter.PublicJsonOptions)));
                }
                else
                {
                    foreach (var entry in entries)
                    {
                        Console.WriteLine($"{entry!["format"]!.GetValue<string>()} {entry["element"]!.GetValue<string>()}");
                    }
                }
                return 0;
            }, json);
        });
        return command;
    }

    private static Command BuildSchemaValidateCommand(Option<bool> jsonOption)
    {
        var formatOption = new Option<string?>("--format")
        {
            Description = "Restrict validation to one format: docx, xlsx, pptx, hwpx, or hwp",
        };
        var command = new Command("validate", "Parse embedded schemas and report load errors");
        command.Add(formatOption);
        command.Add(jsonOption);

        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            return SafeRun(() =>
            {
                var format = result.GetValue(formatOption);
                var formats = string.IsNullOrWhiteSpace(format)
                    ? SchemaHelpLoader.ListFormats()
                    : new[] { SchemaHelpLoader.NormalizeFormat(format!) };
                var errors = new JsonArray();
                var checkedCount = 0;

                foreach (var fmt in formats)
                {
                    foreach (var element in SchemaHelpLoader.ListElements(fmt))
                    {
                        checkedCount++;
                        try
                        {
                            using var _ = SchemaHelpLoader.LoadSchema(fmt, element);
                        }
                        catch (Exception ex)
                        {
                            errors.Add((JsonNode?)new JsonObject
                            {
                                ["format"] = fmt,
                                ["element"] = element,
                                ["message"] = ex.Message,
                            });
                        }
                    }
                }

                if (json)
                {
                    var data = new JsonObject
                    {
                        ["schemaVersion"] = 1,
                        ["checked"] = checkedCount,
                        ["ok"] = errors.Count == 0,
                        ["errors"] = errors,
                    };
                    Console.WriteLine(OutputFormatter.WrapEnvelope(data.ToJsonString(OutputFormatter.PublicJsonOptions)));
                }
                else
                {
                    Console.WriteLine(errors.Count == 0
                        ? $"Schema validation OK ({checkedCount} entries)"
                        : $"Schema validation failed ({errors.Count} errors / {checkedCount} entries)");
                }
                return errors.Count == 0 ? 0 : 1;
            }, json);
        });
        return command;
    }

    private static JsonArray BuildSchemaListEntries(IEnumerable<string> formats)
    {
        var entries = new JsonArray();
        foreach (var fmt in formats)
        {
            foreach (var element in SchemaHelpLoader.ListElements(fmt))
            {
                entries.Add((JsonNode?)new JsonObject
                {
                    ["format"] = fmt,
                    ["element"] = element,
                });
            }
        }
        return entries;
    }

    private static JsonArray BuildFormatArray(IEnumerable<string> formats)
    {
        var array = new JsonArray();
        foreach (var format in formats) array.Add((JsonNode?)JsonValue.Create(format));
        return array;
    }
}
