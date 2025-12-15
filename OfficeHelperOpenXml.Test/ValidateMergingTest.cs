using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using OfficeHelperOpenXml.Api;
using Xunit;
using Xunit.Abstractions;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Validates that text run merging works correctly with real PowerPoint files
    /// Task 8: Validate with real PowerPoint files
    /// Requirements: 4.1, 4.3
    /// </summary>
    public class ValidateMergingTest
    {
        private readonly ITestOutputHelper _output;

        public ValidateMergingTest(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void TestTextboxPptx_GenerateAndInspectJson()
        {
            // Arrange
            string testDir = AppDomain.CurrentDomain.BaseDirectory;
            string solutionRoot = Path.GetFullPath(Path.Combine(testDir, "..", "..", "..", ".."));
            string pptPath = Path.Combine(solutionRoot, "test_ppt", "textbox.pptx");
            string outputPath = Path.Combine(solutionRoot, "test_ppt", "textbox_merged_output.json");
            
            if (!File.Exists(pptPath))
            {
                throw new FileNotFoundException($"Test file not found: {pptPath}");
            }
            
            // Act
            string json;
            using (var reader = new PowerPointReader())
            {
                reader.Load(pptPath);
                json = reader.ToJson();
            }
            
            // Save the output for manual inspection
            File.WriteAllText(outputPath, json);
            _output.WriteLine($"✓ JSON output saved to: {outputPath}");
            _output.WriteLine($"✓ JSON length: {json.Length} characters");
            
            // Parse and inspect the JSON structure
            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                var root = doc.RootElement;
                
                // Log the top-level properties
                _output.WriteLine("\nTop-level JSON properties:");
                foreach (var prop in root.EnumerateObject())
                {
                    _output.WriteLine($"  - {prop.Name}");
                }
                
                // Find and count gradient text runs
                int gradientTextCount = 0;
                int totalTextRuns = 0;
                
                // Check content_slides (not "slides")
                if (root.TryGetProperty("content_slides", out var contentSlides))
                {
                    foreach (var slide in contentSlides.EnumerateArray())
                    {
                        if (slide.TryGetProperty("shapes", out var shapes))
                        {
                            foreach (var shape in shapes.EnumerateArray())
                            {
                                if (shape.TryGetProperty("text", out var textArray))
                                {
                                    foreach (var textRun in textArray.EnumerateArray())
                                    {
                                        totalTextRuns++;
                                        
                                        if (textRun.TryGetProperty("content", out var content))
                                        {
                                            string? contentStr = content.GetString();
                                            if (contentStr != null && contentStr.Contains("文本填充-渐变"))
                                            {
                                                gradientTextCount++;
                                                _output.WriteLine($"\nFound gradient text: '{contentStr}'");
                                                
                                                // Verify it's a complete merged run (not split)
                                                Assert.True(contentStr.StartsWith("文本填充-渐变"), 
                                                    $"Gradient text should start with '文本填充-渐变', got: {contentStr}");
                                                
                                                // Check if it has text_fill
                                                if (textRun.TryGetProperty("text_fill", out var textFill))
                                                {
                                                    _output.WriteLine($"  - Has text_fill property");
                                                    if (textFill.TryGetProperty("fill_type", out var fillType))
                                                    {
                                                        _output.WriteLine($"  - Fill type: {fillType.GetString()}");
                                                        Assert.Equal("gradient", fillType.GetString());
                                                    }
                                                }
                                                else
                                                {
                                                    Assert.Fail($"Gradient text '{contentStr}' should have text_fill property");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                _output.WriteLine($"\n✓ Total text runs: {totalTextRuns}");
                _output.WriteLine($"✓ Gradient text runs found: {gradientTextCount}");
                
                // Assert that we found some gradient text
                Assert.True(gradientTextCount > 0, "Should find at least one gradient text run");
            }
        }
    }
}
