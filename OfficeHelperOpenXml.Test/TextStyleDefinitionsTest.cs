using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Core.Converters;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests to verify that text style definitions are applied consistently
    /// across all slide types (presentation, master slides, layouts, content slides)
    /// </summary>
    public class TextStyleDefinitionsTest
    {
        private readonly string _testJsonPath = Path.Combine("..", "..", "..", "..", "test_ppt", "textbox.json");

        [Fact]
        public void Convert_PresentationHasDefaultTextStyle()
        {
            // Arrange
            var outputPath = Path.GetTempFileName() + ".pptx";
            var converter = new JsonToPptxConverter();

            try
            {
                // Act
                var result = converter.Convert(_testJsonPath, outputPath);

                // Assert
                Assert.True(result, "Conversion should succeed");

                using (var document = PresentationDocument.Open(outputPath, false))
                {
                    var presentation = document.PresentationPart.Presentation;
                    var defaultTextStyle = presentation.Elements<DefaultTextStyle>().FirstOrDefault();

                    Assert.NotNull(defaultTextStyle);
                }
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void Convert_MasterSlidesHaveTextStyles()
        {
            // Arrange
            var outputPath = Path.GetTempFileName() + ".pptx";
            var converter = new JsonToPptxConverter();

            try
            {
                // Act
                var result = converter.Convert(_testJsonPath, outputPath);

                // Assert
                Assert.True(result, "Conversion should succeed");

                using (var document = PresentationDocument.Open(outputPath, false))
                {
                    var masterParts = document.PresentationPart.SlideMasterParts;
                    Assert.NotEmpty(masterParts);

                    foreach (var masterPart in masterParts)
                    {
                        var textStyles = masterPart.SlideMaster.Elements<TextStyles>().FirstOrDefault();
                        Assert.NotNull(textStyles);

                        // Verify all required child styles exist
                        Assert.NotNull(textStyles.Elements<TitleStyle>().FirstOrDefault());
                        Assert.NotNull(textStyles.Elements<BodyStyle>().FirstOrDefault());
                        Assert.NotNull(textStyles.Elements<OtherStyle>().FirstOrDefault());
                    }
                }
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void Convert_AllParagraphsHaveEndParaRPr()
        {
            // Arrange
            var outputPath = Path.GetTempFileName() + ".pptx";
            var converter = new JsonToPptxConverter();

            try
            {
                // Act
                var result = converter.Convert(_testJsonPath, outputPath);

                // Assert
                Assert.True(result, "Conversion should succeed");

                using (var document = PresentationDocument.Open(outputPath, false))
                {
                    // Check content slides
                    var slideParts = document.PresentationPart.SlideParts;
                    foreach (var slidePart in slideParts)
                    {
                        var paragraphs = slidePart.Slide.Descendants<A.Paragraph>();
                        foreach (var paragraph in paragraphs)
                        {
                            var endParaRPr = paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
                            Assert.NotNull(endParaRPr);
                        }
                    }

                    // Check master slides
                    var masterParts = document.PresentationPart.SlideMasterParts;
                    foreach (var masterPart in masterParts)
                    {
                        var paragraphs = masterPart.SlideMaster.Descendants<A.Paragraph>();
                        foreach (var paragraph in paragraphs)
                        {
                            var endParaRPr = paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
                            Assert.NotNull(endParaRPr);
                        }
                    }

                    // Check layouts
                    foreach (var masterPart in masterParts)
                    {
                        var layoutParts = masterPart.SlideLayoutParts;
                        foreach (var layoutPart in layoutParts)
                        {
                            var paragraphs = layoutPart.SlideLayout.Descendants<A.Paragraph>();
                            foreach (var paragraph in paragraphs)
                            {
                                var endParaRPr = paragraph.Elements<A.EndParagraphRunProperties>().FirstOrDefault();
                                Assert.NotNull(endParaRPr);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }
    }
}
