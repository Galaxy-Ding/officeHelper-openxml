using System;
using System.Collections.Generic;
using Xunit;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests to verify backward compatibility of JSON output
    /// Ensures that existing properties are maintained when WordArt styles are added
    /// </summary>
    public class BackwardCompatibilityTests
    {
        /// <summary>
        /// Test that all existing properties are present in JSON output for text runs without WordArt styling
        /// Requirements: 5.1, 5.3
        /// </summary>
        [Fact]
        public void TestBasicTextRun_AllExistingPropertiesPresent()
        {
            // Create a text component with only basic properties (no WordArt styling)
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Basic Text",
                                FontName = "Arial",
                                FontSize = 14.0f,
                                IsBold = true,
                                IsItalic = false,
                                IsUnderline = true,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 255, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify all existing properties are present
            Assert.Contains("\"content\":\"Basic Text\"", json);
            Assert.Contains("\"font\":\"Arial\"", json);
            Assert.Contains("\"font_size\":14.0", json);
            Assert.Contains("\"font_color\":\"RGB(255, 0, 0)\"", json);
            Assert.Contains("\"font_bold\":1", json);
            Assert.Contains("\"font_italic\":0", json);
            Assert.Contains("\"font_underline\":1", json);
            Assert.Contains("\"font_strikethrough\":0", json);
        }

        /// <summary>
        /// Test that existing properties are maintained when WordArt fill is added
        /// Requirements: 5.1
        /// </summary>
        [Fact]
        public void TestTextRunWithFill_ExistingPropertiesMaintained()
        {
            // Create a text component with basic properties AND text fill
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Styled Text",
                                FontName = "Calibri",
                                FontSize = 18.0f,
                                IsBold = false,
                                IsItalic = true,
                                IsUnderline = false,
                                IsStrikethrough = true,
                                FontColor = new ColorInfo { Red = 0, Green = 128, Blue = 255 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Solid,
                                    Color = new ColorInfo { Red = 255, Green = 215, Blue = 0 }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify all existing properties are still present
            Assert.Contains("\"content\":\"Styled Text\"", json);
            Assert.Contains("\"font\":\"Calibri\"", json);
            Assert.Contains("\"font_size\":18.0", json);
            Assert.Contains("\"font_color\":\"RGB(0, 128, 255)\"", json);
            Assert.Contains("\"font_bold\":0", json);
            Assert.Contains("\"font_italic\":1", json);
            Assert.Contains("\"font_underline\":0", json);
            Assert.Contains("\"font_strikethrough\":1", json);
            
            // Verify WordArt fill is also present
            Assert.Contains("\"text_fill\":", json);
        }

        /// <summary>
        /// Test that existing properties are maintained when WordArt outline is added
        /// Requirements: 5.1
        /// </summary>
        [Fact]
        public void TestTextRunWithOutline_ExistingPropertiesMaintained()
        {
            // Create a text component with basic properties AND text outline
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Outlined Text",
                                FontName = "Times New Roman",
                                FontSize = 16.0f,
                                IsBold = true,
                                IsItalic = true,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextOutline = new TextOutlineInfo
                                {
                                    HasOutline = true,
                                    Width = 2.5f,
                                    Color = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                                    DashStyle = LineDashStyle.Solid
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify all existing properties are still present
            Assert.Contains("\"content\":\"Outlined Text\"", json);
            Assert.Contains("\"font\":\"Times New Roman\"", json);
            Assert.Contains("\"font_size\":16.0", json);
            Assert.Contains("\"font_color\":\"RGB(0, 0, 0)\"", json);
            Assert.Contains("\"font_bold\":1", json);
            Assert.Contains("\"font_italic\":1", json);
            Assert.Contains("\"font_underline\":0", json);
            Assert.Contains("\"font_strikethrough\":0", json);
            
            // Verify WordArt outline is also present
            Assert.Contains("\"text_outline\":", json);
        }

        /// <summary>
        /// Test that existing properties are maintained when WordArt effects are added
        /// Requirements: 5.1
        /// </summary>
        [Fact]
        public void TestTextRunWithEffects_ExistingPropertiesMaintained()
        {
            // Create a text component with basic properties AND text effects
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Effect Text",
                                FontName = "Verdana",
                                FontSize = 20.0f,
                                IsBold = false,
                                IsItalic = false,
                                IsUnderline = true,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 128, Green = 0, Blue = 128 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = new ShadowInfo
                                    {
                                        HasShadow = true,
                                        Type = ShadowType.Outer,
                                        Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                        Blur = 3.0f,
                                        Distance = 2.0f,
                                        Angle = 45.0f
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify all existing properties are still present
            Assert.Contains("\"content\":\"Effect Text\"", json);
            Assert.Contains("\"font\":\"Verdana\"", json);
            Assert.Contains("\"font_size\":20.0", json);
            Assert.Contains("\"font_color\":\"RGB(128, 0, 128)\"", json);
            Assert.Contains("\"font_bold\":0", json);
            Assert.Contains("\"font_italic\":0", json);
            Assert.Contains("\"font_underline\":1", json);
            Assert.Contains("\"font_strikethrough\":0", json);
            
            // Verify WordArt effects are also present
            Assert.Contains("\"text_effects\":", json);
        }

        /// <summary>
        /// Test that existing properties are maintained when all WordArt styles are added
        /// Requirements: 5.1
        /// </summary>
        [Fact]
        public void TestTextRunWithAllWordArtStyles_ExistingPropertiesMaintained()
        {
            // Create a text component with basic properties AND all WordArt styles
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Full Style Text",
                                FontName = "Georgia",
                                FontSize = 24.0f,
                                IsBold = true,
                                IsItalic = true,
                                IsUnderline = true,
                                IsStrikethrough = true,
                                FontColor = new ColorInfo { Red = 64, Green = 64, Blue = 64 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Solid,
                                    Color = new ColorInfo { Red = 255, Green = 192, Blue = 203 }
                                },
                                TextOutline = new TextOutlineInfo
                                {
                                    HasOutline = true,
                                    Width = 1.5f,
                                    Color = new ColorInfo { Red = 139, Green = 0, Blue = 139 },
                                    DashStyle = LineDashStyle.Dash
                                },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasGlow = true,
                                    Glow = new GlowInfo
                                    {
                                        Radius = 4.0f,
                                        Color = new ColorInfo { Red = 255, Green = 255, Blue = 0 }
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify all existing properties are still present
            Assert.Contains("\"content\":\"Full Style Text\"", json);
            Assert.Contains("\"font\":\"Georgia\"", json);
            Assert.Contains("\"font_size\":24.0", json);
            Assert.Contains("\"font_color\":\"RGB(64, 64, 64)\"", json);
            Assert.Contains("\"font_bold\":1", json);
            Assert.Contains("\"font_italic\":1", json);
            Assert.Contains("\"font_underline\":1", json);
            Assert.Contains("\"font_strikethrough\":1", json);
            
            // Verify all WordArt styles are also present
            Assert.Contains("\"text_fill\":", json);
            Assert.Contains("\"text_outline\":", json);
            Assert.Contains("\"text_effects\":", json);
        }

        /// <summary>
        /// Test that text runs without WordArt styles output default/null values for WordArt properties
        /// Requirements: 5.3
        /// </summary>
        [Fact]
        public void TestTextRunWithoutWordArt_NoWordArtPropertiesInOutput()
        {
            // Create a text component with only basic properties (no WordArt styling)
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Plain Text",
                                FontName = "Arial",
                                FontSize = 12.0f,
                                IsBold = false,
                                IsItalic = false,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                                // No TextFill, TextOutline, or TextEffects set
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify existing properties are present
            Assert.Contains("\"content\":\"Plain Text\"", json);
            Assert.Contains("\"font\":\"Arial\"", json);
            Assert.Contains("\"font_size\":12.0", json);
            
            // Verify WordArt properties are NOT present (since they're null)
            Assert.DoesNotContain("\"text_fill\":", json);
            Assert.DoesNotContain("\"text_outline\":", json);
            Assert.DoesNotContain("\"text_effects\":", json);
        }

        /// <summary>
        /// Test multiple text runs with mixed styling (some with WordArt, some without)
        /// Requirements: 5.1, 5.3
        /// </summary>
        [Fact]
        public void TestMultipleTextRuns_MixedStyling()
        {
            // Create a text component with multiple runs - some with WordArt, some without
            var textComponent = new TextComponent
            {
                HasText = true,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            // Run 1: Basic properties only
                            new TextRunInfo
                            {
                                Text = "Plain ",
                                FontName = "Arial",
                                FontSize = 12.0f,
                                IsBold = false,
                                IsItalic = false,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            // Run 2: With WordArt fill
                            new TextRunInfo
                            {
                                Text = "Styled ",
                                FontName = "Arial",
                                FontSize = 12.0f,
                                IsBold = true,
                                IsItalic = false,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Solid,
                                    Color = new ColorInfo { Red = 255, Green = 0, Blue = 0 }
                                }
                            },
                            // Run 3: Basic properties only
                            new TextRunInfo
                            {
                                Text = "Text",
                                FontName = "Arial",
                                FontSize = 12.0f,
                                IsBold = false,
                                IsItalic = true,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Verify all runs have their existing properties
            Assert.Contains("\"content\":\"Plain \"", json);
            Assert.Contains("\"content\":\"Styled \"", json);
            Assert.Contains("\"content\":\"Text\"", json);
            
            // Count occurrences of basic properties (should be 3 for each)
            Assert.Equal(3, CountOccurrences(json, "\"font\":\"Arial\""));
            Assert.Equal(3, CountOccurrences(json, "\"font_size\":12.0"));
            Assert.Equal(3, CountOccurrences(json, "\"font_color\":\"RGB(0, 0, 0)\""));
            
            // Verify WordArt fill only appears once (for the styled run)
            Assert.Equal(1, CountOccurrences(json, "\"text_fill\":"));
        }

        /// <summary>
        /// Helper method to count occurrences of a substring in a string
        /// </summary>
        private int CountOccurrences(string text, string pattern)
        {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(pattern, index)) != -1)
            {
                count++;
                index += pattern.Length;
            }
            return count;
        }
    }
}
