// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildDumpCommand(Option<bool> jsonOption)
    {
        var dumpFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx)" };
        var formatOpt = new Option<string>("--format")
        {
            Description = "Output format (currently: batch)",
            DefaultValueFactory = _ => "batch"
        };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Write output to a file instead of stdout" };

        var dumpCommand = new Command("dump", "Serialize a document into a replayable batch script (round-trip mechanism)");
        dumpCommand.Add(dumpFileArg);
        dumpCommand.Add(formatOpt);
        dumpCommand.Add(outOpt);
        dumpCommand.Add(jsonOption);

        dumpCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(dumpFileArg)!;
            var format = (result.GetValue(formatOpt) ?? "batch").ToLowerInvariant();
            var outPath = result.GetValue(outOpt);

            if (format != "batch")
                throw new CliException($"Unsupported --format: {format}. Valid: batch")
                    { Code = "invalid_format", ValidValues = ["batch"] };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".docx")
                throw new CliException($"dump currently supports .docx only (got {ext})")
                    { Code = "unsupported_format" };

            using var word = new WordHandler(file.FullName, editable: false);
            var items = BatchEmitter.EmitWord(word);

            var jsonOpts = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            };
            var output = JsonSerializer.Serialize(items, BatchJsonContext.Default.ListBatchItem);
            if (outPath != null)
            {
                File.WriteAllText(outPath, output);
                Console.WriteLine(outPath);
            }
            else
            {
                Console.WriteLine(output);
            }
            return 0;
        }));

        return dumpCommand;
    }
}
