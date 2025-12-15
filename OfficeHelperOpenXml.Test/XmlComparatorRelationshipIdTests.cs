using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Core.Comparison;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests for XmlComparator's ability to ignore relationship ID differences
    /// </summary>
    public class XmlComparatorRelationshipIdTests : IDisposable
    {
        private readonly XmlComparator _comparator;
        private readonly string _testDir;

        public XmlComparatorRelationshipIdTests()
        {
            _comparator = new XmlComparator();
            _testDir = Path.Combine(Path.GetTempPath(), $"XmlComparatorRelIdTests_{Guid.NewGuid()}");
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
        }

        [Fact]
        public void Compare_DifferentRelationshipIds_ShouldIgnoreDifference()
        {
            // Arrange - Create two .rels files with different relationship IDs
            string xml1 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme1.xml""/>
    <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"" Target=""../slideMasters/slideMaster1.xml""/>
</Relationships>";

            string xml2 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""R8a3f2b1c"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme1.xml""/>
    <Relationship Id=""Rc4d9e7f2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"" Target=""../slideMasters/slideMaster1.xml""/>
</Relationships>";

            string file1 = Path.Combine(_testDir, "test1.xml.rels");
            string file2 = Path.Combine(_testDir, "test2.xml.rels");
            File.WriteAllText(file1, xml1);
            File.WriteAllText(file2, xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert - Should have no differences because relationship IDs are ignored
            Assert.Empty(result.Differences);
        }

        [Fact]
        public void Compare_DifferentEmbedIds_ShouldIgnoreDifference()
        {
            // Arrange - Create two slide files with different r:embed IDs
            string xml1 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:sld xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main"" 
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:cSld>
        <p:spTree>
            <p:pic>
                <p:blipFill>
                    <a:blip xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" r:embed=""rId1""/>
                </p:blipFill>
            </p:pic>
        </p:spTree>
    </p:cSld>
</p:sld>";

            string xml2 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<p:sld xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main"" 
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <p:cSld>
        <p:spTree>
            <p:pic>
                <p:blipFill>
                    <a:blip xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" r:embed=""R8a3f2b1c""/>
                </p:blipFill>
            </p:pic>
        </p:spTree>
    </p:cSld>
</p:sld>";

            string file1 = Path.Combine(_testDir, "slide1.xml");
            string file2 = Path.Combine(_testDir, "slide2.xml");
            File.WriteAllText(file1, xml1);
            File.WriteAllText(file2, xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert - Should have no differences because r:embed IDs are ignored
            Assert.Empty(result.Differences);
        }

        [Fact]
        public void Compare_DifferentNonRelationshipAttributes_ShouldDetectDifference()
        {
            // Arrange - Create two files with different non-relationship attributes
            string xml1 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme1.xml""/>
</Relationships>";

            string xml2 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme2.xml""/>
</Relationships>";

            string file1 = Path.Combine(_testDir, "test1.xml.rels");
            string file2 = Path.Combine(_testDir, "test2.xml.rels");
            File.WriteAllText(file1, xml1);
            File.WriteAllText(file2, xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert - Should detect difference in Target attribute (not a relationship ID)
            Assert.NotEmpty(result.Differences);
            Assert.Contains(result.Differences, d => d.XPath.Contains("@Target"));
        }

        [Fact]
        public void Compare_SameContentDifferentRelIds_ShouldBeIdentical()
        {
            // Arrange - Two files with identical content but different relationship IDs
            string xml1 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme1.xml""/>
    <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"" Target=""../slideLayouts/slideLayout1.xml""/>
    <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image1.png""/>
</Relationships>";

            string xml2 = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""R892a545c540a47c8"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"" Target=""../theme/theme1.xml""/>
    <Relationship Id=""R2149c202145e4705"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"" Target=""../slideLayouts/slideLayout1.xml""/>
    <Relationship Id=""Rf3a8b9c4d5e6f7a8"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image1.png""/>
</Relationships>";

            string file1 = Path.Combine(_testDir, "test1.xml.rels");
            string file2 = Path.Combine(_testDir, "test2.xml.rels");
            File.WriteAllText(file1, xml1);
            File.WriteAllText(file2, xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert - Should be considered identical
            Assert.Empty(result.Differences);
        }
    }
}
