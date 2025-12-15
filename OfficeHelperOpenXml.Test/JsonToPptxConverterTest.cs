using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Core.Converters;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeHelperOpenXml.Test
{
    public class JsonToPptxConverterTest
    {
        [Fact]
        public void Convert_WithValidJson_CreatesValidPptx()
        {
            // Arrange
            var workspaceRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
            var jsonPath = Path.Combine(workspaceRoot, "test_ppt", "textbox.json");
            var outputPath = Path.Combine(workspaceRoot, "test_ppt", "textbox_from_converter.pptx");
            var converter = new JsonToPptxConverter();

            // Clean up any existing output file
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Act
            var result = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");
            Assert.True(File.Exists(outputPath), "Output PPTX file should exist");

            // Verify the PPTX can be opened
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document);
                Assert.NotNull(document.PresentationPart);
                Assert.NotNull(document.PresentationPart.Presentation);
                
                var slideIdList = document.PresentationPart.Presentation.SlideIdList;
                Assert.NotNull(slideIdList);
                Assert.True(slideIdList.ChildElements.Count > 0, "Should have at least one slide");
                
                Console.WriteLine($"✅ Created PPTX with {slideIdList.ChildElements.Count} slides");
            }
        }

        [Fact]
        public void Convert_WithNonExistentJson_ReturnsFalse()
        {
            // Arrange
            var jsonPath = "nonexistent.json";
            var outputPath = "test_output.pptx";
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.False(result, "Conversion should fail for non-existent file");
        }

        [Fact]
        public void Convert_WithEmptyJsonPath_ReturnsFalse()
        {
            // Arrange
            var jsonPath = "";
            var outputPath = "test_output.pptx";
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.False(result, "Conversion should fail for empty JSON path");
        }

        [Fact]
        public void Convert_WithEmptyOutputPath_ReturnsFalse()
        {
            // Arrange
            var workspaceRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
            var jsonPath = Path.Combine(workspaceRoot, "test_ppt", "textbox.json");
            var outputPath = "";
            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.False(result, "Conversion should fail for empty output path");
        }
    }
}
