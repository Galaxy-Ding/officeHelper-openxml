using System;
using System.IO;
using System.Linq;
using OfficeHelperOpenXml.Core.Comparison;
using Xunit;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for XmlComparator component
    /// </summary>
    public class XmlComparatorTests : IDisposable
    {
        private readonly string _testDir;
        private readonly XmlComparator _comparator;

        public XmlComparatorTests()
        {
            _comparator = new XmlComparator();
            _testDir = Path.Combine(Path.GetTempPath(), $"XmlComparatorTests_{Guid.NewGuid()}");
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            // Cleanup test directory
            if (Directory.Exists(_testDir))
            {
                try
                {
                    Directory.Delete(_testDir, true);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        [Fact]
        public void Compare_IdenticalXml_ShouldReturnNoDifferences()
        {
            // Arrange
            string xml = @"<?xml version=""1.0""?>
<root>
    <element attr=""value"">Content</element>
</root>";
            
            string file1 = CreateTempXmlFile("identical1.xml", xml);
            string file2 = CreateTempXmlFile("identical2.xml", xml);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.NotNull(result);
            Assert.True(result.AreIdentical);
            Assert.Empty(result.Differences);
        }

        [Fact]
        public void Compare_DifferentAttributeValue_ShouldDetectDifference()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element attr=""value1"">Content</element>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element attr=""value2"">Content</element>
</root>";
            
            string file1 = CreateTempXmlFile("attr1.xml", xml1);
            string file2 = CreateTempXmlFile("attr2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.NotEmpty(result.Differences);
            Assert.Contains(result.Differences, d => d.Type == DifferenceType.DifferentAttributeValue);
        }

        [Fact]
        public void Compare_MissingAttribute_ShouldDetectDifference()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element attr=""value"">Content</element>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element>Content</element>
</root>";
            
            string file1 = CreateTempXmlFile("withattr.xml", xml1);
            string file2 = CreateTempXmlFile("noattr.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.Contains(result.Differences, d => d.Type == DifferenceType.MissingAttribute);
        }

        [Fact]
        public void Compare_ExtraAttribute_ShouldDetectDifference()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element>Content</element>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element attr=""value"">Content</element>
</root>";
            
            string file1 = CreateTempXmlFile("noattr2.xml", xml1);
            string file2 = CreateTempXmlFile("withattr2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.Contains(result.Differences, d => d.Type == DifferenceType.ExtraAttribute);
        }

        [Fact]
        public void Compare_MissingElement_ShouldDetectDifference()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element1>Content</element1>
    <element2>Content</element2>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element1>Content</element1>
</root>";
            
            string file1 = CreateTempXmlFile("twoelem.xml", xml1);
            string file2 = CreateTempXmlFile("oneelem.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.Contains(result.Differences, d => d.Type == DifferenceType.MissingElement);
        }

        [Fact]
        public void Compare_ExtraElement_ShouldDetectDifference()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element1>Content</element1>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element1>Content</element1>
    <element2>Content</element2>
</root>";
            
            string file1 = CreateTempXmlFile("oneelem2.xml", xml1);
            string file2 = CreateTempXmlFile("twoelem2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.Contains(result.Differences, d => d.Type == DifferenceType.ExtraElement);
        }

        [Fact]
        public void Compare_DifferentTextContent_ShouldDetectDifference()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element>Content1</element>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element>Content2</element>
</root>";
            
            string file1 = CreateTempXmlFile("text1.xml", xml1);
            string file2 = CreateTempXmlFile("text2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.Contains(result.Differences, d => d.Type == DifferenceType.DifferentTextContent);
        }

        [Fact]
        public void Compare_WhitespaceNormalization_ShouldIgnoreWhitespaceDifferences()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element>Content   with   spaces</element>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element>Content with spaces</element>
</root>";
            
            string file1 = CreateTempXmlFile("spaces1.xml", xml1);
            string file2 = CreateTempXmlFile("spaces2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.True(result.AreIdentical, "Whitespace differences should be normalized");
        }

        [Fact]
        public void Compare_XPathGeneration_ShouldIncludeXPathForAllDifferences()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <element attr=""value1"">Content1</element>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <element attr=""value2"">Content2</element>
</root>";
            
            string file1 = CreateTempXmlFile("xpath1.xml", xml1);
            string file2 = CreateTempXmlFile("xpath2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.All(result.Differences, d => Assert.False(string.IsNullOrEmpty(d.XPath)));
        }

        [Fact]
        public void Compare_InvalidXml_ShouldThrowComparisonException()
        {
            // Arrange
            string invalidXml = "This is not valid XML";
            string validXml = @"<?xml version=""1.0""?><root></root>";
            
            string file1 = CreateTempXmlFile("invalid.xml", invalidXml);
            string file2 = CreateTempXmlFile("valid.xml", validXml);

            // Act & Assert
            var exception = Assert.Throws<ComparisonException>(() => _comparator.Compare(file1, file2));
            Assert.Equal(ErrorCategory.Xml, exception.Category);
        }

        [Fact]
        public void Compare_NestedElements_ShouldCompareRecursively()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <parent>
        <child attr=""value1"">Content</child>
    </parent>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <parent>
        <child attr=""value2"">Content</child>
    </parent>
</root>";
            
            string file1 = CreateTempXmlFile("nested1.xml", xml1);
            string file2 = CreateTempXmlFile("nested2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            Assert.Contains(result.Differences, d => 
                d.Type == DifferenceType.DifferentAttributeValue && 
                d.XPath.Contains("child"));
        }

        [Fact]
        public void Compare_MultipleElementsWithSameName_ShouldUseIndexing()
        {
            // Arrange
            string xml1 = @"<?xml version=""1.0""?>
<root>
    <item>First</item>
    <item>Second</item>
    <item>Third</item>
</root>";
            
            string xml2 = @"<?xml version=""1.0""?>
<root>
    <item>First</item>
    <item>Modified</item>
    <item>Third</item>
</root>";
            
            string file1 = CreateTempXmlFile("multi1.xml", xml1);
            string file2 = CreateTempXmlFile("multi2.xml", xml2);

            // Act
            var result = _comparator.Compare(file1, file2);

            // Assert
            Assert.False(result.AreIdentical);
            var diff = result.Differences.FirstOrDefault(d => d.Type == DifferenceType.DifferentTextContent);
            Assert.NotNull(diff);
            Assert.Contains("[2]", diff.XPath); // Should reference second item
        }

        private string CreateTempXmlFile(string filename, string content)
        {
            string path = Path.Combine(_testDir, filename);
            File.WriteAllText(path, content);
            return path;
        }
    }
}
