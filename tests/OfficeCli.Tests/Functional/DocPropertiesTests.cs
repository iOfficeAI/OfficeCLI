// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for document properties (title, author, subject, etc.) on Excel and PowerPoint handlers.
/// </summary>
public class DocPropertiesTests
{
    // ==================== Excel Document Properties ====================

    public class ExcelDocProperties : IDisposable
    {
        private readonly string _path;
        private ExcelHandler _handler;

        public ExcelDocProperties()
        {
            _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
            BlankDocCreator.Create(_path);
            _handler = new ExcelHandler(_path, editable: true);
        }

        public void Dispose()
        {
            _handler.Dispose();
            if (File.Exists(_path)) File.Delete(_path);
        }

        private ExcelHandler Reopen()
        {
            _handler.Dispose();
            _handler = new ExcelHandler(_path, editable: true);
            return _handler;
        }

        [Fact]
        public void Set_Title_IsReadBack()
        {
            _handler.Set("/", new() { ["title"] = "My Workbook" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("title");
            node.Format["title"].Should().Be("My Workbook");
        }

        [Fact]
        public void Set_Author_IsReadBack()
        {
            _handler.Set("/", new() { ["author"] = "John Doe" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("author");
            node.Format["author"].Should().Be("John Doe");
        }

        [Fact]
        public void Set_Subject_IsReadBack()
        {
            _handler.Set("/", new() { ["subject"] = "Test Subject" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("subject");
            node.Format["subject"].Should().Be("Test Subject");
        }

        [Fact]
        public void Set_Description_IsReadBack()
        {
            _handler.Set("/", new() { ["description"] = "A test workbook" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("description");
            node.Format["description"].Should().Be("A test workbook");
        }

        [Fact]
        public void Set_Category_IsReadBack()
        {
            _handler.Set("/", new() { ["category"] = "Testing" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("category");
            node.Format["category"].Should().Be("Testing");
        }

        [Fact]
        public void Set_Keywords_IsReadBack()
        {
            _handler.Set("/", new() { ["keywords"] = "test, excel, properties" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("keywords");
            node.Format["keywords"].Should().Be("test, excel, properties");
        }

        [Fact]
        public void Set_LastModifiedBy_IsReadBack()
        {
            _handler.Set("/", new() { ["lastModifiedBy"] = "Jane Smith" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("lastModifiedBy");
            node.Format["lastModifiedBy"].Should().Be("Jane Smith");
        }

        [Fact]
        public void Set_Revision_IsReadBack()
        {
            _handler.Set("/", new() { ["revision"] = "5" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("revision");
            node.Format["revision"].Should().Be("5");
        }

        [Fact]
        public void Set_MultipleProperties_AllReadBack()
        {
            _handler.Set("/", new()
            {
                ["title"] = "Budget 2025",
                ["author"] = "Finance Team",
                ["subject"] = "Annual Budget",
                ["category"] = "Finance"
            });

            var node = _handler.Get("/");
            node.Format["title"].Should().Be("Budget 2025");
            node.Format["author"].Should().Be("Finance Team");
            node.Format["subject"].Should().Be("Annual Budget");
            node.Format["category"].Should().Be("Finance");
        }

        [Fact]
        public void Set_Creator_Alias_IsReadBack_AsAuthor()
        {
            _handler.Set("/", new() { ["creator"] = "Bob" });
            var node = _handler.Get("/");
            node.Format["author"].Should().Be("Bob");
        }

        [Fact]
        public void Set_Properties_PersistAfterReopen()
        {
            _handler.Set("/", new()
            {
                ["title"] = "Persistent Title",
                ["author"] = "Persistent Author"
            });

            Reopen();

            var node = _handler.Get("/");
            node.Format["title"].Should().Be("Persistent Title");
            node.Format["author"].Should().Be("Persistent Author");
        }

        [Fact]
        public void Set_UnknownProperty_ReturnsUnsupported()
        {
            var unsupported = _handler.Set("/", new() { ["foobar"] = "baz" });
            unsupported.Should().Contain("foobar");
        }
    }

    // ==================== PowerPoint Document Properties ====================

    public class PptxDocProperties : IDisposable
    {
        private readonly string _path;
        private PowerPointHandler _handler;

        public PptxDocProperties()
        {
            _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
            BlankDocCreator.Create(_path);
            _handler = new PowerPointHandler(_path, editable: true);
        }

        public void Dispose()
        {
            _handler.Dispose();
            if (File.Exists(_path)) File.Delete(_path);
        }

        private PowerPointHandler Reopen()
        {
            _handler.Dispose();
            _handler = new PowerPointHandler(_path, editable: true);
            return _handler;
        }

        [Fact]
        public void Set_Title_IsReadBack()
        {
            _handler.Set("/", new() { ["title"] = "My Presentation" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("title");
            node.Format["title"].Should().Be("My Presentation");
        }

        [Fact]
        public void Set_Author_IsReadBack()
        {
            _handler.Set("/", new() { ["author"] = "John Doe" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("author");
            node.Format["author"].Should().Be("John Doe");
        }

        [Fact]
        public void Set_Subject_IsReadBack()
        {
            _handler.Set("/", new() { ["subject"] = "Quarterly Review" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("subject");
            node.Format["subject"].Should().Be("Quarterly Review");
        }

        [Fact]
        public void Set_Description_IsReadBack()
        {
            _handler.Set("/", new() { ["description"] = "Q4 presentation" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("description");
            node.Format["description"].Should().Be("Q4 presentation");
        }

        [Fact]
        public void Set_Category_IsReadBack()
        {
            _handler.Set("/", new() { ["category"] = "Business" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("category");
            node.Format["category"].Should().Be("Business");
        }

        [Fact]
        public void Set_Keywords_IsReadBack()
        {
            _handler.Set("/", new() { ["keywords"] = "pptx, test" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("keywords");
            node.Format["keywords"].Should().Be("pptx, test");
        }

        [Fact]
        public void Set_LastModifiedBy_IsReadBack()
        {
            _handler.Set("/", new() { ["lastModifiedBy"] = "Jane" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("lastModifiedBy");
            node.Format["lastModifiedBy"].Should().Be("Jane");
        }

        [Fact]
        public void Set_Revision_IsReadBack()
        {
            _handler.Set("/", new() { ["revision"] = "3" });
            var node = _handler.Get("/");
            node.Format.Should().ContainKey("revision");
            node.Format["revision"].Should().Be("3");
        }

        [Fact]
        public void Set_MultipleProperties_AllReadBack()
        {
            _handler.Set("/", new()
            {
                ["title"] = "Strategy Deck",
                ["author"] = "Leadership",
                ["keywords"] = "strategy, planning"
            });

            var node = _handler.Get("/");
            node.Format["title"].Should().Be("Strategy Deck");
            node.Format["author"].Should().Be("Leadership");
            node.Format["keywords"].Should().Be("strategy, planning");
        }

        [Fact]
        public void Set_Properties_CoexistWithSlideSize()
        {
            _handler.Set("/", new()
            {
                ["title"] = "Mixed Props",
                ["slideSize"] = "16:9"
            });

            var node = _handler.Get("/");
            node.Format["title"].Should().Be("Mixed Props");
            node.Format["slideSize"].Should().Be("widescreen");
        }

        [Fact]
        public void Set_Properties_PersistAfterReopen()
        {
            _handler.Set("/", new()
            {
                ["title"] = "Persistent Title",
                ["author"] = "Persistent Author"
            });

            Reopen();

            var node = _handler.Get("/");
            node.Format["title"].Should().Be("Persistent Title");
            node.Format["author"].Should().Be("Persistent Author");
        }
    }
}
