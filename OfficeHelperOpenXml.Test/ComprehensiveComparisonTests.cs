using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Core.Comparison;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Comprehensive comparison tests for validating zero differences after all fixes
    /// Tests with textbox.json (original test file), simple presentation, and complex presentation
    /// Validates Requirements: 16.3
    /// </summary>
    public class ComprehensiveComparisonTests
    {
        private static readonly string TestDataDir = Path.Combine("..", "..", "..", "..", "test_ppt");
        private static readonly string OutputDir = Path.Combine("..", "..", "..", "..", "test_comparison_output");

        public ComprehensiveComparisonTests()
        {
            // Ensure output directory exists
            if (!Directory.Exists(OutputDir))
            {
                Directory.CreateDirectory(OutputDir);
            }
        }

        [Fact]
        public void Test_TextboxJson_ZeroDifferences()
        {
            // Arrange
            var jsonPath = Path.Combine(TestDataDir, "textbox.json");
            var originalPath = Path.Combine(TestDataDir, "textbox.pptx");
            var generatedPath = Path.Combine(OutputDir, "textbox_generated.pptx");
            var comparisonOutputDir = Path.Combine(OutputDir, "textbox_comparison");

            // Skip if test files don't exist
            if (!File.Exists(jsonPath) || !File.Exists(originalPath))
            {
                // Create output to indicate test was skipped
                File.WriteAllText(Path.Combine(OutputDir, "textbox_test_skipped.txt"), 
                    $"Test skipped: Required files not found\nJSON: {jsonPath}\nOriginal: {originalPath}");
                return;
            }

            // Act - Convert JSON to PPTX
            var converter = new JsonToPptxConverter();
            bool conversionSuccess = converter.Convert(jsonPath, generatedPath);

            // Assert conversion succeeded
            Assert.True(conversionSuccess, "JSON to PPTX conversion should succeed");
            Assert.True(File.Exists(generatedPath), "Generated PPTX file should exist");

            // Act - Run comparison
            var config = new ComparisonConfig
            {
                ContentOriginalPath = originalPath,
                ContentGeneratedPath = generatedPath,
                OutputDirectory = comparisonOutputDir
            };

            var orchestrator = new ComparisonOrchestrator();
            var summary = orchestrator.ExecuteComparison(config);

            // Assert - Zero differences
            Assert.NotNull(summary);
            Assert.NotNull(summary.ContentComparison);
            Assert.Equal(0, summary.ContentComparison.TotalDifferences);
        }

        [Fact]
        public void Test_SimplePresentation_ZeroDifferences()
        {
            // Arrange - Create simple JSON
            var simpleJson = @"{
  ""master_slides"": [],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Simple Test"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""TextBox 1"",
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
              ""content"": ""Hello World!"",
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

            var jsonPath = Path.Combine(OutputDir, "simple_test.json");
            var generatedPath = Path.Combine(OutputDir, "simple_generated.pptx");
            var regeneratedPath = Path.Combine(OutputDir, "simple_regenerated.pptx");
            var comparisonOutputDir = Path.Combine(OutputDir, "simple_comparison");

            File.WriteAllText(jsonPath, simpleJson);

            // Act - Convert JSON to PPTX twice
            var converter = new JsonToPptxConverter();
            bool firstConversion = converter.Convert(jsonPath, generatedPath);
            bool secondConversion = converter.Convert(jsonPath, regeneratedPath);

            // Assert conversions succeeded
            Assert.True(firstConversion, "First conversion should succeed");
            Assert.True(secondConversion, "Second conversion should succeed");

            // Act - Compare the two generated files (should be identical)
            var config = new ComparisonConfig
            {
                ContentOriginalPath = generatedPath,
                ContentGeneratedPath = regeneratedPath,
                OutputDirectory = comparisonOutputDir
            };

            var orchestrator = new ComparisonOrchestrator();
            var summary = orchestrator.ExecuteComparison(config);

            // Assert - Zero differences (same JSON should produce identical PPTX)
            Assert.NotNull(summary);
            Assert.NotNull(summary.ContentComparison);
            Assert.Equal(0, summary.ContentComparison.TotalDifferences);
        }

        [Fact]
        public void Test_ComplexPresentation_ZeroDifferences()
        {
            // Arrange - Create complex JSON with multiple masters and slides
            var complexJson = @"{
  ""master_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Master 1"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""Title Placeholder"",
          ""box"": ""2.00,1.00,20.00,3.00"",
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
              ""content"": ""Master Title"",
              ""font"": ""Arial"",
              ""font_size"": 24.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 1,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ],
  ""content_slides"": [
    {
      ""page_number"": 1,
      ""title"": ""Slide 1"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""TextBox 1"",
          ""box"": ""5.00,5.00,15.00,5.00"",
          ""rotation"": 0.0,
          ""fill"": {
            ""color"": ""RGB(200, 200, 255)"",
            ""opacity"": 1.0
          },
          ""line"": {
            ""has_outline"": 1,
            ""color"": ""RGB(0, 0, 255)"",
            ""width"": 2.0
          },
          ""shadow"": {
            ""has_shadow"": 1,
            ""color"": ""RGB(100, 100, 100)"",
            ""opacity"": 0.5,
            ""blur"": 4.0,
            ""distance"": 3.0,
            ""angle"": 45.0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": ""Bold Text"",
              ""font"": ""Arial"",
              ""font_size"": 18.0,
              ""font_color"": ""RGB(0, 0, 0)"",
              ""font_bold"": 1,
              ""font_italic"": 0,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    },
    {
      ""page_number"": 2,
      ""title"": ""Slide 2"",
      ""sub_title"": """",
      ""shapes"": [
        {
          ""type"": ""textbox"",
          ""name"": ""TextBox 2"",
          ""box"": ""3.00,3.00,18.00,8.00"",
          ""rotation"": 15.0,
          ""fill"": {
            ""color"": ""RGB(255, 200, 200)"",
            ""opacity"": 0.8
          },
          ""line"": {
            ""has_outline"": 1,
            ""color"": ""RGB(255, 0, 0)"",
            ""width"": 1.5
          },
          ""shadow"": {
            ""has_shadow"": 0
          },
          ""hastext"": 1,
          ""text"": [
            {
              ""content"": ""Italic Text"",
              ""font"": ""Calibri"",
              ""font_size"": 16.0,
              ""font_color"": ""RGB(128, 0, 0)"",
              ""font_bold"": 0,
              ""font_italic"": 1,
              ""font_underline"": 0,
              ""font_strikethrough"": 0
            }
          ]
        }
      ]
    }
  ]
}";

            var jsonPath = Path.Combine(OutputDir, "complex_test.json");
            var generatedPath = Path.Combine(OutputDir, "complex_generated.pptx");
            var regeneratedPath = Path.Combine(OutputDir, "complex_regenerated.pptx");
            var comparisonOutputDir = Path.Combine(OutputDir, "complex_comparison");

            File.WriteAllText(jsonPath, complexJson);

            // Act - Convert JSON to PPTX twice
            var converter = new JsonToPptxConverter();
            bool firstConversion = converter.Convert(jsonPath, generatedPath);
            bool secondConversion = converter.Convert(jsonPath, regeneratedPath);

            // Assert conversions succeeded
            Assert.True(firstConversion, "First conversion should succeed");
            Assert.True(secondConversion, "Second conversion should succeed");

            // Act - Compare the two generated files (should be identical)
            var config = new ComparisonConfig
            {
                ContentOriginalPath = generatedPath,
                ContentGeneratedPath = regeneratedPath,
                OutputDirectory = comparisonOutputDir
            };

            var orchestrator = new ComparisonOrchestrator();
            var summary = orchestrator.ExecuteComparison(config);

            // Assert - Zero differences
            Assert.NotNull(summary);
            Assert.NotNull(summary.ContentComparison);
            Assert.Equal(0, summary.ContentComparison.TotalDifferences);
        }
    }
}
