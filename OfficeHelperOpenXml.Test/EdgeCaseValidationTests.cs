using System;
using System.IO;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Edge case validation tests for PPTX generator
    /// Tests edge cases like empty presentations, shapes with zero dimensions, etc.
    /// Validates Requirements: 16.3
    /// </summary>
    public class EdgeCaseValidationTests
    {
        private static readonly string OutputDir = Path.Combine("..", "..", "..", "..", "test_comparison_output", "edge_cases");

        public EdgeCaseValidationTests()
        {
            // Ensure output directory exists
            if (!Directory.Exists(OutputDir))
            {
                Directory.CreateDirectory(OutputDir);
            }
        }

        [Fact]
        public void Test_PresentationWithOnlyMasterSlides()
        {
            // Arrange - JSON with only master slides, no content slides
            var json = @"{
  ""master_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Master Only"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Master TextBox"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(255, 255, 255)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": ""Master Text"",
              ""font"": ""Arial"",
              ""font_size"": 18.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 0,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ],
  ""content_slides"": []
}";

            var jsonPath = Path.Combine(OutputDir, "master_only.json");
            var outputPath = Path.Combine(OutputDir, "master_only.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with only master slides");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify structure
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                Assert.NotNull(document.PresentationPart.Presentation);
                
                // Should have master slides but no content slides
                var masterSlides = document.PresentationPart.SlideMasterParts;
                Assert.NotEmpty(masterSlides);
                
                var contentSlides = document.PresentationPart.SlideParts;
                Assert.Empty(contentSlides);
            }
        }

        [Fact]
        public void Test_PresentationWithOnlyContentSlides()
        {
            // Arrange - JSON with only content slides, no master slides
            var json = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Content Only"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Content TextBox"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(255, 255, 255)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": ""Content Text"",
              ""font"": ""Arial"",
              ""font_size"": 18.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 0,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "content_only.json");
            var outputPath = Path.Combine(OutputDir, "content_only.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with only content slides");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify structure
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                Assert.NotNull(document.PresentationPart.Presentation);
                
                // Should have at least one master slide (default) and content slides
                var masterSlides = document.PresentationPart.SlideMasterParts;
                Assert.NotEmpty(masterSlides);
                
                var contentSlides = document.PresentationPart.SlideParts;
                Assert.NotEmpty(contentSlides);
            }
        }

        [Fact]
        public void Test_ShapeWithZeroDimensions()
        {
            // Arrange - Shape with zero width and height
            var json = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Zero Dimensions"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Zero Size Shape"",
          ""box"": ""5.00,5.00,0.00,0.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(255, 255, 255)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 0,
          ""text"": []
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "zero_dimensions.json");
            var outputPath = Path.Combine(OutputDir, "zero_dimensions.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with zero dimension shapes");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify file can be opened
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                var slides = document.PresentationPart.SlideParts;
                Assert.NotEmpty(slides);
            }
        }

        [Fact]
        public void Test_ShapeWithNoText()
        {
            // Arrange - Shape with hastext=0 and empty text array
            var json = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""No Text"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Empty TextBox"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(200, 200, 200)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 1,
            ""color"": ""RGB(0, 0, 0)"",
            ""width"": 1.0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 0,
          ""text"": []
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "no_text.json");
            var outputPath = Path.Combine(OutputDir, "no_text.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with shapes having no text");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify file can be opened
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                var slides = document.PresentationPart.SlideParts;
                Assert.NotEmpty(slides);
            }
        }

        [Fact]
        public void Test_ShapeWithNoFill()
        {
            // Arrange - Shape with no fill (transparent)
            var json = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""No Fill"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""No Fill Shape"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(255, 255, 255)"",
            ""opacity"": 0.0
          },
          ""line"": {
            ""has_outline"": 1,
            ""color"": ""RGB(0, 0, 0)"",
            ""width"": 2.0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": ""No Fill"",
              ""font"": ""Arial"",
              ""font_size"": 18.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 0,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "no_fill.json");
            var outputPath = Path.Combine(OutputDir, "no_fill.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with shapes having no fill");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify file can be opened
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                var slides = document.PresentationPart.SlideParts;
                Assert.NotEmpty(slides);
            }
        }

        [Fact]
        public void Test_ShapeWithNoLineAndNoShadow()
        {
            // Arrange - Shape with no outline and no shadow
            var json = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""No Line No Shadow"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Minimal Shape"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(200, 200, 255)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": ""Minimal"",
              ""font"": ""Arial"",
              ""font_size"": 18.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 0,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "no_line_no_shadow.json");
            var outputPath = Path.Combine(OutputDir, "no_line_no_shadow.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with minimal shape properties");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify file can be opened
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                var slides = document.PresentationPart.SlideParts;
                Assert.NotEmpty(slides);
            }
        }

        [Fact]
        public void Test_EmptyTextRun()
        {
            // Arrange - Text run with empty content
            var json = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Empty Text Run"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Empty Text"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(255, 255, 255)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 1,
            ""color"": ""RGB(0, 0, 0)"",
            ""width"": 1.0
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": """",
              ""font"": ""Arial"",
              ""font_size"": 18.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 0,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "empty_text_run.json");
            var outputPath = Path.Combine(OutputDir, "empty_text_run.pptx");
            File.WriteAllText(jsonPath, json);

            // Act
            var converter = new JsonToPptxConverter();
            bool success = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(success, "Conversion should succeed with empty text runs");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify file can be opened
            using (var document = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(document.PresentationPart);
                var slides = document.PresentationPart.SlideParts;
                Assert.NotEmpty(slides);
            }
        }
    }
}
