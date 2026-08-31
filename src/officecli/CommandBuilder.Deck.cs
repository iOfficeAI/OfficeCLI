// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Deck;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildDeckCommand(Option<bool> rootJsonOption)
    {
        var deck = new Command("deck", "Build and validate WorkMate semantic presentation decks");

        var catalog = new Command("catalog", "Print the embedded themes and semantic layouts");
        catalog.Add(rootJsonOption);
        catalog.SetAction(result => RunDeck(() =>
        {
            Console.WriteLine(JsonSerializer.Serialize(DeckCatalogLoader.Load(), DeckJsonContext.Default.DeckCatalog));
            return 0;
        }));
        deck.Add(catalog);

        var specArg = new Argument<FileInfo>("spec") { Description = "Path to a *.workmate-deck.json file" };
        var validate = new Command("validate", "Validate a WorkMate DeckSpec without changing files");
        validate.Add(specArg);
        validate.Add(rootJsonOption);
        validate.SetAction(result => RunDeck(() =>
        {
            var specFile = result.GetValue(specArg)!;
            var spec = DeckService.LoadSpec(specFile.FullName);
            var validation = DeckService.Validate(spec, specFile.FullName);
            Console.WriteLine(JsonSerializer.Serialize(validation, DeckJsonContext.Default.DeckValidationResult));
            return validation.Valid ? 0 : 1;
        }));
        deck.Add(validate);

        var buildSpecArg = new Argument<FileInfo>("spec") { Description = "Path to a *.workmate-deck.json file" };
        var outputOption = new Option<FileInfo>("--output") { Description = "Output .pptx path", Required = true };
        var expectedRevisionOption = new Option<long?>("--expected-revision") { Description = "Reject stale DeckSpec builds" };
        outputOption.Aliases.Add("-o");
        var build = new Command("build", "Compile a WorkMate DeckSpec into an editable PPTX");
        build.Add(buildSpecArg);
        build.Add(outputOption);
        build.Add(expectedRevisionOption);
        build.Add(rootJsonOption);
        build.SetAction(result => RunDeck(() =>
        {
            var specFile = result.GetValue(buildSpecArg)!;
            var spec = DeckService.LoadSpec(specFile.FullName);
            var expectedRevision = result.GetValue(expectedRevisionOption);
            var output = DeckService.Build(spec, specFile.FullName, result.GetValue(outputOption)!.FullName, expectedRevision);
            Console.WriteLine(JsonSerializer.Serialize(
                new DeckBuildResult(true, output, spec.Revision),
                DeckJsonContext.Default.DeckBuildResult));
            return 0;
        }));
        deck.Add(build);

        var renderSpecArg = new Argument<FileInfo>("spec") { Description = "Path to a *.workmate-deck.json file" };
        var formatOption = new Option<string>("--format") { DefaultValueFactory = _ => "preview" };
        var render = new Command("render", "Render normalized preview scene data");
        render.Add(renderSpecArg);
        render.Add(formatOption);
        render.Add(rootJsonOption);
        render.SetAction(result => RunDeck(() =>
        {
            var format = result.GetValue(formatOption);
            if (!string.Equals(format, "preview", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("deck render currently supports only --format preview.");
            var specFile = result.GetValue(renderSpecArg)!;
            var spec = DeckService.LoadSpec(specFile.FullName);
            Console.WriteLine(JsonSerializer.Serialize(
                DeckService.RenderPreview(spec, specFile.FullName),
                DeckJsonContext.Default.DeckPreviewScene));
            return 0;
        }));
        deck.Add(render);

        return deck;
    }

    private static int RunDeck(Func<int> action) => SafeRun(() =>
    {
        try
        {
            return action();
        }
        catch (Exception exception)
        {
            if (Environment.GetEnvironmentVariable("OFFICECLI_DECK_DEBUG") == "1")
                Console.Error.WriteLine(exception);
            throw;
        }
    }, json: true);
}
