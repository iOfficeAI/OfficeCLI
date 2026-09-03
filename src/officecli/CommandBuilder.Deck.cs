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

        var layoutQuery = new Command("layout-query", "Rank catalog layouts for a slide by role and content capacity hints");
        var roleOption = new Option<string?>("--role") { Description = "Semantic role filter (cover, metrics, trend, …)" };
        var itemCountOption = new Option<int?>("--item-count") { Description = "Approximate module/item/KPI count the slide needs" };
        itemCountOption.Aliases.Add("--module-count");
        var hasChartOption = new Option<bool?>("--has-chart") { Description = "Whether the slide needs a chart-capable slot (true/false)" };
        var needsMediaOption = new Option<bool?>("--needs-media") { Description = "Whether the slide needs an image/media slot (true/false)" };
        needsMediaOption.Aliases.Add("--has-image");
        var hasTableOption = new Option<bool?>("--has-table") { Description = "Whether the slide needs a table-capable slot (true/false)" };
        var queryOption = new Option<string?>("--query") { Description = "Optional free-text hint matched against layout id/label/role" };
        var limitOption = new Option<int>("--limit") { Description = "Max ranked results (1-50)", DefaultValueFactory = _ => 8 };
        layoutQuery.Add(roleOption);
        layoutQuery.Add(itemCountOption);
        layoutQuery.Add(hasChartOption);
        layoutQuery.Add(needsMediaOption);
        layoutQuery.Add(hasTableOption);
        layoutQuery.Add(queryOption);
        layoutQuery.Add(limitOption);
        layoutQuery.Add(rootJsonOption);
        layoutQuery.SetAction(result => RunDeck(() =>
        {
            var request = new DeckLayoutQueryRequest(
                Role: result.GetValue(roleOption),
                ItemCount: result.GetValue(itemCountOption),
                HasChart: result.GetValue(hasChartOption),
                NeedsMedia: result.GetValue(needsMediaOption),
                HasTable: result.GetValue(hasTableOption),
                Query: result.GetValue(queryOption),
                Limit: result.GetValue(limitOption));
            var response = DeckLayoutQuery.Query(request);
            Console.WriteLine(JsonSerializer.Serialize(response, DeckJsonContext.Default.DeckLayoutQueryResult));
            return 0;
        }));
        deck.Add(layoutQuery);


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


        var remapSpecArg = new Argument<FileInfo>("spec") { Description = "Path to a *.workmate-deck.json file" };
        var toThemeOption = new Option<string>("--to") { Description = "Target catalog theme.id", Required = true };
        toThemeOption.Aliases.Add("--theme");
        var applyOption = new Option<bool>("--apply") { Description = "Write theme.id (+ optional report) back to the spec", DefaultValueFactory = _ => false };
        var writeReportOption = new Option<bool>("--write-report") { Description = "Embed report under extensions.themeRemap when applying", DefaultValueFactory = _ => true };
        var setModeOption = new Option<bool>("--set-mode") { Description = "Set theme.mode from target token luminance", DefaultValueFactory = _ => true };
        var remapLimitOption = new Option<int>("--limit") { Description = "Max same-role alternatives per slide (1-20)", DefaultValueFactory = _ => 5 };
        var outputSpecOption = new Option<FileInfo?>("--output") { Description = "Optional output path when applying (default: in-place)" };
        outputSpecOption.Aliases.Add("-o");
        var themeRemap = new Command("theme-remap", "Change theme.id, report layouts that may need remap, suggest same-role alternatives via layout-query");
        themeRemap.Add(remapSpecArg);
        themeRemap.Add(toThemeOption);
        themeRemap.Add(applyOption);
        themeRemap.Add(writeReportOption);
        themeRemap.Add(setModeOption);
        themeRemap.Add(remapLimitOption);
        themeRemap.Add(outputSpecOption);
        themeRemap.Add(rootJsonOption);
        themeRemap.SetAction(result => RunDeck(() =>
        {
            var specFile = result.GetValue(remapSpecArg)!;
            var spec = DeckService.LoadSpec(specFile.FullName);
            var apply = result.GetValue(applyOption);
            var output = result.GetValue(outputSpecOption);
            var remap = DeckThemeRemap.Remap(
                spec,
                result.GetValue(toThemeOption)!,
                new DeckThemeRemapOptions(
                    Apply: apply,
                    WriteReport: result.GetValue(writeReportOption),
                    SetMode: result.GetValue(setModeOption),
                    Limit: result.GetValue(remapLimitOption),
                    SpecPath: specFile.FullName));
            if (apply)
            {
                var targetPath = output?.FullName ?? specFile.FullName;
                DeckThemeRemap.SaveSpec(remap.Spec!, targetPath);
                remap = remap with { SpecPath = targetPath, Spec = null };
            }
            else
            {
                remap = remap with { Spec = null };
            }
            Console.WriteLine(JsonSerializer.Serialize(remap, DeckJsonContext.Default.DeckThemeRemapResult));
            return 0;
        }));
        deck.Add(themeRemap);

        var scaffold = new Command("scaffold", "Outline a long WorkMate deck (section transitions + role mix) from goal/audience/page count");
        var goalOption = new Option<string?>("--goal") { Description = "Deck goal / one-line intent" };
        var audienceOption = new Option<string?>("--audience") { Description = "Primary audience" };
        var pagesOption = new Option<int>("--pages") { Description = "Target slide count (4-60)", DefaultValueFactory = _ => 12 };
        pagesOption.Aliases.Add("--page-count");
        var titleOption = new Option<string?>("--title") { Description = "Deck title (defaults from goal)" };
        var languageOption = new Option<string>("--language") { Description = "BCP-47 language tag", DefaultValueFactory = _ => "en-US" };
        languageOption.Aliases.Add("--lang");
        var scaffoldThemeOption = new Option<string>("--theme") { Description = "Catalog theme.id", DefaultValueFactory = _ => "csbu-workmate" };
        var seedOption = new Option<string?>("--seed") { Description = "Reproducibility seed (default: hash of goal|audience|pages)" };
        var scaffoldOutputOption = new Option<FileInfo>("--output") { Description = "Output *.workmate-deck.json path", Required = true };
        scaffoldOutputOption.Aliases.Add("-o");
        var writeSpecOption = new Option<bool>("--write") { Description = "Write the outline DeckSpec to --output (default true)", DefaultValueFactory = _ => true };
        scaffold.Add(goalOption);
        scaffold.Add(audienceOption);
        scaffold.Add(pagesOption);
        scaffold.Add(titleOption);
        scaffold.Add(languageOption);
        scaffold.Add(scaffoldThemeOption);
        scaffold.Add(seedOption);
        scaffold.Add(scaffoldOutputOption);
        scaffold.Add(writeSpecOption);
        scaffold.Add(rootJsonOption);
        scaffold.SetAction(result => RunDeck(() =>
        {
            var request = new DeckScaffoldRequest(
                Goal: result.GetValue(goalOption),
                Audience: result.GetValue(audienceOption),
                Pages: result.GetValue(pagesOption),
                Title: result.GetValue(titleOption),
                Language: result.GetValue(languageOption),
                ThemeId: result.GetValue(scaffoldThemeOption),
                Seed: result.GetValue(seedOption));
            var built = DeckScaffold.Scaffold(request);
            var output = result.GetValue(scaffoldOutputOption)!;
            if (result.GetValue(writeSpecOption))
                DeckScaffold.SaveSpec(built.Spec, output.FullName);
            var payload = new DeckScaffoldCliResult(
                Written: result.GetValue(writeSpecOption),
                SpecPath: output.FullName,
                SlideCount: built.Spec.Slides.Count,
                Stage: built.Spec.Stage,
                Report: built.Report);
            Console.WriteLine(JsonSerializer.Serialize(payload, DeckJsonContext.Default.DeckScaffoldCliResult));
            return 0;
        }));
        deck.Add(scaffold);


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
