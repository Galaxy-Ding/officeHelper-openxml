using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Xunit;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Integration test for Phase 2 file generation
    /// Verifies that core.xml and layout .rels files are generated
    /// </summary>
    public class Phase2FileGenerationIntegrationTest : IDisposable
    {
        private readonly string _testJsonPath;
        private readonly string _testOutputPath;

        public Phase2FileGenerationIntegrationTest()
        {
            // Get the test project directory and navigate to the workspace root
            var testDir = Directory.GetCurrentDirectory();
            var workspaceRoot = Path.GetFullPath(Path.Combine(testDir, "..", "..", "..", ".."));
            _testJsonPath = Path.Combine(workspaceRoot, "test_ppt", "textbox.json");
            _testOutputPath = Path.Combine(Path.GetTempPath(), $"test_phase2_{Guid.NewGuid()}.pptx");
        }

        public void Dispose()
        {
            if (File.Exists(_testOutputPath))
            {
                try
                {
                    File.Delete(_testOutputPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        [Fact]
        public void Convert_GeneratesCoreXmlFile()
        {
            // Arrange
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(_testJsonPath, _testOutputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");
            Assert.True(File.Exists(_testOutputPath), "Output file should exist");

            // Verify core.xml exists in the package
            using (var archive = ZipFile.OpenRead(_testOutputPath))
            {
                var coreXmlEntry = archive.Entries.FirstOrDefault(e => e.FullName == "docProps/core.xml");
                Assert.NotNull(coreXmlEntry);
            }
        }

        [Fact]
        public void Convert_GeneratesLayoutRelationshipFiles()
        {
            // Arrange
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(_testJsonPath, _testOutputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");
            Assert.True(File.Exists(_testOutputPath), "Output file should exist");

            // Verify layout .rels files exist in the package
            using (var archive = ZipFile.OpenRead(_testOutputPath))
            {
                var layoutRelsEntries = archive.Entries
                    .Where(e => e.FullName.StartsWith("ppt/slideLayouts/_rels/") && e.FullName.EndsWith(".rels"))
                    .ToList();

                Assert.NotEmpty(layoutRelsEntries);
            }
        }

        [Fact]
        public void Convert_CoreXmlContainsRequiredProperties()
        {
            // Arrange
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(_testJsonPath, _testOutputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");

            // Read and verify core.xml content
            using (var archive = ZipFile.OpenRead(_testOutputPath))
            {
                var coreXmlEntry = archive.Entries.FirstOrDefault(e => e.FullName == "docProps/core.xml");
                Assert.NotNull(coreXmlEntry);

                using (var stream = coreXmlEntry.Open())
                using (var reader = new StreamReader(stream))
                {
                    var content = reader.ReadToEnd();

                    // Verify required properties are present
                    Assert.Contains("OfficeHelperOpenXml", content); // Creator
                    Assert.Contains("<cp:revision>1</cp:revision>", content); // Revision
                }
            }
        }

        [Fact]
        public void Convert_LayoutRelsContainsRelationshipToMaster()
        {
            // Arrange
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(_testJsonPath, _testOutputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");

            // Read and verify layout .rels content
            using (var archive = ZipFile.OpenRead(_testOutputPath))
            {
                var layoutRelsEntry = archive.Entries
                    .FirstOrDefault(e => e.FullName.StartsWith("ppt/slideLayouts/_rels/") && e.FullName.EndsWith(".rels"));

                Assert.NotNull(layoutRelsEntry);

                using (var stream = layoutRelsEntry.Open())
                using (var reader = new StreamReader(stream))
                {
                    var content = reader.ReadToEnd();

                    // Verify relationship to master slide exists
                    Assert.Contains("slideMaster", content);
                    Assert.Contains("rId", content); // Sequential relationship ID
                }
            }
        }

        [Fact]
        public void Convert_AllRequiredFilesExist()
        {
            // Arrange
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(_testJsonPath, _testOutputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");

            // Verify all required files exist
            using (var archive = ZipFile.OpenRead(_testOutputPath))
            {
                var fileNames = archive.Entries.Select(e => e.FullName).ToList();

                // Core metadata files
                Assert.Contains("docProps/core.xml", fileNames);
                Assert.Contains("docProps/app.xml", fileNames);

                // Presentation files
                Assert.Contains("ppt/presentation.xml", fileNames);
                Assert.Contains("ppt/presProps.xml", fileNames);
                Assert.Contains("ppt/viewProps.xml", fileNames);
                Assert.Contains("ppt/tableStyles.xml", fileNames);

                // Layout relationship files
                var layoutRelsFiles = fileNames.Where(f => f.StartsWith("ppt/slideLayouts/_rels/") && f.EndsWith(".rels")).ToList();
                Assert.NotEmpty(layoutRelsFiles);
            }
        }
    }
}
