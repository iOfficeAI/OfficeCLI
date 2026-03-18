using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BatchFunctionalTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private readonly string _docxPath;

    public BatchFunctionalTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
    }

    public void Dispose()
    {
        foreach (var p in new[] { _xlsxPath, _pptxPath, _docxPath })
            if (File.Exists(p)) File.Delete(p);
    }

    [Fact]
    public void Batch_Excel_AddMultipleCells_SingleOpenSave()
    {
        BlankDocCreator.Create(_xlsxPath);
        using var handler = new ExcelHandler(_xlsxPath, editable: true);

        var items = new List<BatchItem>
        {
            new() { Command = "set", Path = "/Sheet1/A1", Props = new() { ["value"] = "Name" } },
            new() { Command = "set", Path = "/Sheet1/B1", Props = new() { ["value"] = "Age" } },
            new() { Command = "set", Path = "/Sheet1/A2", Props = new() { ["value"] = "Alice" } },
            new() { Command = "set", Path = "/Sheet1/B2", Props = new() { ["value"] = "30" } },
            new() { Command = "set", Path = "/Sheet1/A1", Props = new() { ["bold"] = "true" } },
            new() { Command = "set", Path = "/Sheet1/B1", Props = new() { ["bold"] = "true" } },
        };

        foreach (var item in items)
        {
            var props = item.Props ?? new();
            handler.Set(item.Path!, props);
        }

        // Verify all cells
        var a1 = handler.Get("/Sheet1/A1");
        a1.Text.Should().Be("Name");
        a1.Format.Should().ContainKey("font.bold");

        var b2 = handler.Get("/Sheet1/B2");
        b2.Text.Should().Be("30");

        var a2 = handler.Get("/Sheet1/A2");
        a2.Text.Should().Be("Alice");
    }

    [Fact]
    public void Batch_Pptx_AddSlideAndShapes()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/", Type = "slide", Props = new() { ["title"] = "Slide 1" } },
            new() { Command = "add", Parent = "/slide[1]", Type = "shape", Props = new() { ["text"] = "Hello", ["fill"] = "FF0000" } },
            new() { Command = "add", Parent = "/slide[1]", Type = "shape", Props = new() { ["text"] = "World", ["fill"] = "00FF00" } },
            new() { Command = "set", Path = "/slide[1]/shape[2]", Props = new() { ["bold"] = "true" } },
        };

        foreach (var item in items)
        {
            var props = item.Props ?? new();
            if (item.Command == "add")
                handler.Add(item.Parent!, item.Type!, item.Index, props);
            else if (item.Command == "set")
                handler.Set(item.Path!, props);
        }

        // shape[1] is the title placeholder; added shapes follow
        var shape1 = handler.Get("/slide[1]/shape[2]");
        shape1.Text.Should().Be("Hello");
        shape1.Format["fill"].Should().Be("FF0000");

        var shape2 = handler.Get("/slide[1]/shape[3]");
        shape2.Text.Should().Be("World");
    }

    [Fact]
    public void Batch_Word_AddParagraphsAndFormat()
    {
        BlankDocCreator.Create(_docxPath);
        using var handler = new WordHandler(_docxPath, editable: true);

        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/body", Type = "paragraph", Props = new() { ["text"] = "Title", ["bold"] = "true", ["size"] = "28" } },
            new() { Command = "add", Parent = "/body", Type = "paragraph", Props = new() { ["text"] = "First paragraph" } },
            new() { Command = "add", Parent = "/body", Type = "paragraph", Props = new() { ["text"] = "Second paragraph" } },
            new() { Command = "set", Path = "/body/p[2]", Props = new() { ["italic"] = "true" } },
        };

        foreach (var item in items)
        {
            var props = item.Props ?? new();
            if (item.Command == "add")
                handler.Add(item.Parent!, item.Type!, item.Index, props);
            else if (item.Command == "set")
                handler.Set(item.Path!, props);
        }

        var p1 = handler.Get("/body/p[1]");
        p1.Text.Should().Be("Title");

        var p2 = handler.Get("/body/p[2]");
        p2.Text.Should().Be("First paragraph");
    }

    [Fact]
    public void Batch_MixedOperations_AddSetRemoveGet()
    {
        BlankDocCreator.Create(_pptxPath);
        using var handler = new PowerPointHandler(_pptxPath, editable: true);

        // Add slide and shapes (shape[1] = title placeholder from layout)
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Keep" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Remove me" });

        // Batch: set one, remove another
        handler.Set("/slide[1]/shape[2]", new() { ["bold"] = "true" });
        handler.Remove("/slide[1]/shape[3]");

        var shapes = handler.Query("shape");
        // shape[1] = title, shape[2] = "Keep"
        shapes.Should().HaveCount(2);
        shapes[1].Text.Should().Be("Keep");
    }

    [Fact]
    public void BatchItem_ToResidentRequest_MapsCorrectly()
    {
        var item = new BatchItem
        {
            Command = "set",
            Path = "/slide[1]/shape[1]",
            Props = new() { ["bold"] = "true", ["fill"] = "FF0000" }
        };

        var req = item.ToResidentRequest();

        req.Command.Should().Be("set");
        req.Args["path"].Should().Be("/slide[1]/shape[1]");
        req.Props.Should().Contain("bold=true");
        req.Props.Should().Contain("fill=FF0000");
    }

    [Fact]
    public void BatchItem_ToResidentRequest_HandlesAddWithAllFields()
    {
        var item = new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "paragraph",
            Index = 2,
            Props = new() { ["text"] = "Hello" }
        };

        var req = item.ToResidentRequest();

        req.Command.Should().Be("add");
        req.Args["parent"].Should().Be("/body");
        req.Args["type"].Should().Be("paragraph");
        req.Args["index"].Should().Be("2");
        req.Props.Should().Contain("text=Hello");
    }

    [Fact]
    public void BatchItem_Serialization_RoundTrip()
    {
        var items = new List<BatchItem>
        {
            new() { Command = "add", Parent = "/Sheet1", Type = "cell", Props = new() { ["address"] = "A1", ["value"] = "Test" } },
            new() { Command = "set", Path = "/Sheet1/A1", Props = new() { ["bold"] = "true" } },
        };

        var json = System.Text.Json.JsonSerializer.Serialize(items);
        var deserialized = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(json);

        deserialized.Should().HaveCount(2);
        deserialized![0].Command.Should().Be("add");
        deserialized[0].Props!["address"].Should().Be("A1");
        deserialized[1].Command.Should().Be("set");
        deserialized[1].Path.Should().Be("/Sheet1/A1");
    }
}
