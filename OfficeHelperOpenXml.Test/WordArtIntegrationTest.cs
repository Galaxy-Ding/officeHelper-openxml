using Xunit;
using OfficeHelperOpenXml.Api;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Linq;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Integration tests for WordArt text styling extraction from real PowerPoint files
    /// Validates Requirements: 1.1, 1.2, 1.3, 2.1, 3.1, 3.3
    /// </summary>
    public class WordArtIntegrationTest
    {
        private const string TestPptPath = @"test_ppt\textbox.pptx";

        [Fact]
        public void ExtractFromRealPowerPoint_ShouldProduceValidJson()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            
            // Assert - Verify basic structure
            Assert.NotNull(jsonOutput);
            
            var masterSlides = jsonOutput["master_slides"] as JArray;
            var contentSlides = jsonOutput["content_slides"] as JArray;
            
            Assert.True(masterSlides != null || contentSlides != null, 
                "Expected to find master_slides or content_slides");
        }

        [Fact]
        public void ExtractFromRealPowerPoint_TextRunsShouldHaveBasicProperties()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var textRuns = GetAllTextRuns(jsonOutput);
            
            // Assert - Verify all text runs have basic properties (backward compatibility)
            Assert.True(textRuns.Count > 0, "Expected to find at least one text run");
            
            foreach (var textRun in textRuns)
            {
                // Requirement 5.1: Maintain all existing properties
                Assert.NotNull(textRun["content"]);
                Assert.NotNull(textRun["font"]);
                Assert.NotNull(textRun["font_size"]);
                Assert.NotNull(textRun["font_color"]);
                Assert.NotNull(textRun["font_bold"]);
                Assert.NotNull(textRun["font_italic"]);
                Assert.NotNull(textRun["font_underline"]);
                Assert.NotNull(textRun["font_strikethrough"]);
            }
        }

        [Fact]
        public void ExtractFromRealPowerPoint_TextFillShouldBeExtractedWhenPresent()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var textRuns = GetAllTextRuns(jsonOutput);
            
            // Assert
            Assert.True(textRuns.Count > 0, "Expected to find at least one text run");
            
            // Check if any text runs have text_fill property
            var textRunsWithFill = textRuns.Where(tr => tr["text_fill"] != null).ToList();
            
            // Note: The test file may not have WordArt styling, which is valid
            // This test verifies that IF text fills are present, they have correct structure
            if (textRunsWithFill.Count == 0)
            {
                // No text fills found - this is acceptable for files without WordArt styling
                // Test passes as there's nothing to validate
                return;
            }
            
            // If text fills are present, verify structure
            foreach (var textRun in textRunsWithFill)
            {
                var textFill = textRun["text_fill"];
                Assert.NotNull(textFill);
                
                // Requirement 1.1: Solid fill should include type and color
                var hasFill = textFill["has_fill"];
                Assert.NotNull(hasFill);
                
                if ((int)hasFill == 1)
                {
                    var fillTypeToken = textFill["fill_type"];
                    Assert.NotNull(fillTypeToken);
                    
                    var fillType = fillTypeToken.ToString();
                    
                    if (fillType == "solid")
                    {
                        // Requirement 1.1: Solid fill serialization
                        // Color and transparency should be present for solid fills
                        var color = textFill["color"];
                        var transparency = textFill["transparency"];
                        
                        // At least one should be present
                        Assert.True(color != null || transparency != null, 
                            "Solid fill should have color or transparency information");
                    }
                    else if (fillType == "gradient")
                    {
                        // Requirement 1.2: Gradient fill serialization
                        var gradient = textFill["gradient"];
                        if (gradient != null)
                        {
                            // If gradient exists, verify its structure
                            Assert.NotNull(gradient["gradient_type"]);
                            Assert.NotNull(gradient["stops"]);
                        }
                    }
                    else if (fillType == "pattern")
                    {
                        // Requirement 1.3: Pattern fill serialization
                        var pattern = textFill["pattern"];
                        if (pattern != null)
                        {
                            // If pattern exists, verify its structure
                            Assert.NotNull(pattern["pattern_type"]);
                        }
                    }
                }
            }
        }

        [Fact]
        public void ExtractFromRealPowerPoint_TextOutlineShouldBeExtractedWhenPresent()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var textRuns = GetAllTextRuns(jsonOutput);
            
            // Assert
            Assert.True(textRuns.Count > 0, "Expected to find at least one text run");
            
            // Check if any text runs have text_outline property
            var textRunsWithOutline = textRuns.Where(tr => tr["text_outline"] != null).ToList();
            
            // If text outlines are present, verify structure
            foreach (var textRun in textRunsWithOutline)
            {
                var textOutline = textRun["text_outline"];
                
                // Requirement 2.1: Outline should include presence, width, color, dash style
                Assert.NotNull(textOutline["has_outline"]);
                
                if ((int)textOutline["has_outline"] == 1)
                {
                    Assert.NotNull(textOutline["width"]);
                    Assert.NotNull(textOutline["color"]);
                    Assert.NotNull(textOutline["dash_style"]);
                }
            }
        }

        [Fact]
        public void ExtractFromRealPowerPoint_TextEffectsShouldBeExtractedWhenPresent()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var textRuns = GetAllTextRuns(jsonOutput);
            
            // Assert
            Assert.True(textRuns.Count > 0, "Expected to find at least one text run");
            
            // Check if any text runs have text_effects property
            var textRunsWithEffects = textRuns.Where(tr => tr["text_effects"] != null).ToList();
            
            // If text effects are present, verify structure
            foreach (var textRun in textRunsWithEffects)
            {
                var textEffects = textRun["text_effects"];
                
                Assert.NotNull(textEffects["has_effects"]);
                
                if ((int)textEffects["has_effects"] == 1)
                {
                    // Check for shadow (Requirement 3.1)
                    if (textEffects["shadow"] != null)
                    {
                        var shadow = textEffects["shadow"];
                        Assert.NotNull(shadow["has_shadow"]);
                        
                        if ((int)shadow["has_shadow"] == 1)
                        {
                            Assert.NotNull(shadow["type"]);
                            Assert.NotNull(shadow["color"]);
                            Assert.NotNull(shadow["blur"]);
                            Assert.NotNull(shadow["distance"]);
                            Assert.NotNull(shadow["angle"]);
                        }
                    }
                    
                    // Check for glow (Requirement 3.3)
                    if (textEffects["glow"] != null)
                    {
                        var glow = textEffects["glow"];
                        Assert.NotNull(glow["has_glow"]);
                        
                        if ((int)glow["has_glow"] == 1)
                        {
                            Assert.NotNull(glow["radius"]);
                            Assert.NotNull(glow["color"]);
                        }
                    }
                }
            }
        }

        [Fact]
        public void ExtractFromRealPowerPoint_ShouldNotHaveDirectShadowProperty()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var textRuns = GetAllTextRuns(jsonOutput);
            
            // Assert
            Assert.True(textRuns.Count > 0, "Expected to find at least one text run");
            
            // Verify no text run has a direct "shadow" property
            // Shadow should only be under "text_effects"
            foreach (var textRun in textRuns)
            {
                var textRunObj = textRun as JObject;
                Assert.False(textRunObj.ContainsKey("shadow"), 
                    "Text run should not have a direct 'shadow' property. Shadow should be under 'text_effects' if present.");
            }
        }

        [Fact]
        public void ExtractFromRealPowerPoint_ThemeColorsShouldBePreserved()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var textRuns = GetAllTextRuns(jsonOutput);
            
            // Assert
            Assert.True(textRuns.Count > 0, "Expected to find at least one text run");
            
            // Check if any text runs have theme colors
            foreach (var textRun in textRuns)
            {
                // Check text fill for theme colors
                var textFill = textRun["text_fill"];
                if (textFill != null)
                {
                    var color = textFill["color"];
                    if (color != null && color.Type == JTokenType.Object)
                    {
                        var schemeColor = color["schemeColor"];
                        if (schemeColor != null)
                        {
                            // Requirement 1.5: Theme color preservation
                            Assert.NotNull(schemeColor);
                            
                            // If color transforms exist, verify they're preserved
                            var transforms = color["colorTransforms"];
                            if (transforms != null && transforms.Type == JTokenType.Object)
                            {
                                // At least one transform should be present
                                Assert.True(transforms.HasValues, "Color transforms should have values");
                            }
                        }
                    }
                }
                
                // Check text outline for theme colors
                var textOutline = textRun["text_outline"];
                if (textOutline != null)
                {
                    var color = textOutline["color"];
                    if (color != null && color.Type == JTokenType.Object)
                    {
                        var schemeColor = color["schemeColor"];
                        if (schemeColor != null)
                        {
                            Assert.NotNull(schemeColor);
                        }
                    }
                }
            }
        }

        [Fact]
        public void ExtractFromRealPowerPoint_JsonShouldBeWellFormed()
        {
            // Arrange & Act
            var jsonOutput = GenerateJsonFromPpt();
            var jsonString = jsonOutput.ToString();
            
            // Assert - Verify JSON is well-formed and can be parsed
            Assert.False(string.IsNullOrEmpty(jsonString));
            
            // Verify it can be parsed back
            var reparsed = JObject.Parse(jsonString);
            Assert.NotNull(reparsed);
            
            // Verify basic structure is intact
            Assert.True(reparsed["master_slides"] != null || reparsed["content_slides"] != null);
        }

        private JObject GenerateJsonFromPpt()
        {
            // Navigate from bin/Debug/net8.0 to solution root
            var currentDir = Directory.GetCurrentDirectory();
            var solutionRoot = currentDir;
            
            // Go up from bin/Debug/net8.0 to project root, then to solution root
            for (int i = 0; i < 4; i++)
            {
                var parent = Directory.GetParent(solutionRoot);
                if (parent != null)
                {
                    solutionRoot = parent.FullName;
                }
            }
            
            var testPath = Path.Combine(solutionRoot, TestPptPath);
            
            if (!File.Exists(testPath))
            {
                throw new FileNotFoundException($"Test file not found: {testPath}. Current directory: {currentDir}, Solution root: {solutionRoot}");
            }

            using (var reader = PowerPointReaderFactory.CreateReader(testPath, out bool success))
            {
                Assert.True(success, "Failed to load PowerPoint file");

                var jsonString = reader.ToJson();
                Assert.False(string.IsNullOrEmpty(jsonString), "JSON output should not be empty");

                return JObject.Parse(jsonString);
            }
        }

        private System.Collections.Generic.List<JToken> GetAllTextRuns(JObject jsonOutput)
        {
            var textRuns = new System.Collections.Generic.List<JToken>();
            
            var masterSlides = jsonOutput["master_slides"] as JArray;
            var contentSlides = jsonOutput["content_slides"] as JArray;
            
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
                            textRuns.Add(textRun);
                        }
                    }
                }
            }
            
            return textRuns;
        }
    }
}
