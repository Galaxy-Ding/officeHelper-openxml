using Xunit;
using OfficeHelperOpenXml.Api;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Linq;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for verifying text run shadow property removal
    /// Validates Requirements: 1.1, 1.2, 1.4
    /// </summary>
    public class TextRunShadowRemovalTests
    {
        private const string TestPptPath = @"D:\pythonf\c_sharp_project\officeHelperOpenxml\test_ppt\textbox.pptx";

        [Fact]
        public void TextRun_ShouldNotContainShadowProperty()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            
            // Check both master_slides and content_slides
            var masterSlides = jsonOutput["master_slides"] as JArray;
            var contentSlides = jsonOutput["content_slides"] as JArray;

            // Assert
            Assert.True(masterSlides != null || contentSlides != null, "Expected to find master_slides or content_slides");

            int textRunsChecked = 0;

            // Check all text runs across all slides
            foreach (var slideArray in new[] { masterSlides, contentSlides })
            {
                if (slideArray == null) continue;

                foreach (var slide in slideArray)
                {
                    var shapes = slide["shapes"] as JArray;
                    if (shapes == null) continue;

                    foreach (var shape in shapes)
                    {
                        var textArray = shape["text"] as JArray;
                        if (textArray == null || !textArray.Any()) continue;

                        foreach (var textRun in textArray)
                        {
                            textRunsChecked++;
                            
                            // Verify text run does NOT have shadow property
                            var shadowProperty = textRun["shadow"];
                            Assert.Null(shadowProperty);

                            // Verify text run has expected properties
                            Assert.NotNull(textRun["content"]);
                            Assert.NotNull(textRun["font"]);
                            Assert.NotNull(textRun["font_size"]);
                            Assert.NotNull(textRun["font_color"]);
                        }
                    }
                }
            }

            Assert.True(textRunsChecked > 0, "Expected to find at least one text run");
        }

        [Fact]
        public void Shape_ShouldContainShadowAtShapeLevel()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var masterSlides = jsonOutput["master_slides"] as JArray;
            var contentSlides = jsonOutput["content_slides"] as JArray;

            // Assert
            Assert.True(masterSlides != null || contentSlides != null, "Expected to find master_slides or content_slides");

            bool foundShapeWithShadow = false;
            int shapesChecked = 0;

            // Check that shapes have shadow at shape level
            foreach (var slideArray in new[] { masterSlides, contentSlides })
            {
                if (slideArray == null) continue;

                foreach (var slide in slideArray)
                {
                    var shapes = slide["shapes"] as JArray;
                    if (shapes == null) continue;

                    foreach (var shape in shapes)
                    {
                        shapesChecked++;
                        
                        // Verify shape has shadow property at shape level
                        var shapeShadow = shape["shadow"];
                        Assert.NotNull(shapeShadow);

                        // Verify shadow has expected structure
                        Assert.NotNull(shapeShadow["has_shadow"]);
                        Assert.NotNull(shapeShadow["color"]);
                        Assert.NotNull(shapeShadow["opacity"]);

                        if ((int?)shapeShadow["has_shadow"] == 1)
                        {
                            foundShapeWithShadow = true;
                        }
                    }
                }
            }

            Assert.True(shapesChecked > 0, "Expected to find at least one shape");
            // Verify we found at least one shape with shadow enabled
            Assert.True(foundShapeWithShadow, "Expected to find at least one shape with shadow enabled");
        }

        [Fact]
        public void MultipleTextRuns_ShouldHaveConsistentStructure()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var masterSlides = jsonOutput["master_slides"] as JArray;
            var contentSlides = jsonOutput["content_slides"] as JArray;

            // Assert
            Assert.True(masterSlides != null || contentSlides != null, "Expected to find master_slides or content_slides");

            var textRunProperties = new System.Collections.Generic.HashSet<string>();
            int textRunCount = 0;

            // Collect all property names from all text runs
            foreach (var slideArray in new[] { masterSlides, contentSlides })
            {
                if (slideArray == null) continue;

                foreach (var slide in slideArray)
                {
                    var shapes = slide["shapes"] as JArray;
                    if (shapes == null) continue;

                    foreach (var shape in shapes)
                    {
                        var textArray = shape["text"] as JArray;
                        if (textArray == null || !textArray.Any()) continue;

                        foreach (var textRun in textArray)
                        {
                            textRunCount++;
                            var properties = ((JObject)textRun).Properties().Select(p => p.Name);
                            
                            if (textRunProperties.Count == 0)
                            {
                                // First text run - establish baseline
                                foreach (var prop in properties)
                                {
                                    textRunProperties.Add(prop);
                                }
                            }
                            else
                            {
                                // Subsequent text runs - verify consistency
                                // All text runs should have the same set of properties (excluding optional ones)
                                var currentProps = new System.Collections.Generic.HashSet<string>(properties);
                                
                                // Core properties that must be present
                                Assert.Contains("content", currentProps);
                                Assert.Contains("font", currentProps);
                                Assert.Contains("font_size", currentProps);
                                Assert.Contains("font_color", currentProps);
                                Assert.Contains("font_bold", currentProps);
                                Assert.Contains("font_italic", currentProps);
                                Assert.Contains("font_underline", currentProps);
                                Assert.Contains("font_strikethrough", currentProps);
                                
                                // Shadow should NOT be present
                                Assert.DoesNotContain("shadow", currentProps);
                            }
                        }
                    }
                }
            }

            Assert.True(textRunCount > 0, "Expected to find at least one text run");
        }

        [Fact]
        public void JsonOutput_ShouldNotContainShadowStringInTextRuns()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var masterSlides = jsonOutput["master_slides"] as JArray;
            var contentSlides = jsonOutput["content_slides"] as JArray;

            // Assert
            Assert.True(masterSlides != null || contentSlides != null, "Expected to find master_slides or content_slides");

            int textRunsChecked = 0;

            // Get all text run sections
            foreach (var slideArray in new[] { masterSlides, contentSlides })
            {
                if (slideArray == null) continue;

                foreach (var slide in slideArray)
                {
                    var shapes = slide["shapes"] as JArray;
                    if (shapes == null) continue;

                    foreach (var shape in shapes)
                    {
                        var textArray = shape["text"] as JArray;
                        if (textArray == null || !textArray.Any()) continue;

                        foreach (var textRun in textArray)
                        {
                            textRunsChecked++;
                            var textRunObj = textRun as JObject;
                            
                            // Verify the text run does not have a direct "shadow" property
                            // Shadow should only be nested under "text_effects" if present
                            Assert.False(textRunObj.ContainsKey("shadow"), 
                                "Text run should not have a direct 'shadow' property. Shadow should be under 'text_effects' if present.");
                            
                            // If text_effects exists, shadow can be there (that's the new structure)
                            // But we're specifically checking that shadow is NOT a direct property of the text run
                        }
                    }
                }
            }

            Assert.True(textRunsChecked > 0, "Expected to find at least one text run");
        }

        private JObject GenerateJsonFromPpt()
        {
            if (!File.Exists(TestPptPath))
            {
                throw new FileNotFoundException($"Test file not found: {TestPptPath}");
            }

            using (var reader = PowerPointReaderFactory.CreateReader(TestPptPath, out bool success))
            {
                Assert.True(success, "Failed to load PowerPoint file");

                var jsonString = reader.ToJson();
                Assert.False(string.IsNullOrEmpty(jsonString), "JSON output should not be empty");

                return JObject.Parse(jsonString);
            }
        }
    }
}
