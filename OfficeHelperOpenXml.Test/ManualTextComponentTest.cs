using System;
using System.Collections.Generic;
using Xunit;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Manual test to verify the JSON output format
    /// </summary>
    public class ManualTextComponentTest
    {
        [Fact]
        public void TestTextComponentJsonOutput()
        {
            // Create a simple text component with one paragraph and one run
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
                                Text = "Hello World",
                                FontName = "Arial",
                                FontSize = 12,
                                IsBold = true,
                                IsItalic = false,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                HasShadow = false,
                                Shadow = new ShadowInfo()
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("JSON Output:");
            Console.WriteLine(json);
            
            // Verify that the JSON doesn't contain "shadow"
            Assert.DoesNotContain("\"shadow\"", json);
            
            // Verify that it contains the expected properties
            Assert.Contains("\"hastext\":1", json);
            Assert.Contains("\"text\":[", json);
            Assert.Contains("\"content\":\"Hello World\"", json);
            Assert.Contains("\"font\":\"Arial\"", json);
            Assert.Contains("\"font_size\":12", json);
            Assert.Contains("\"font_bold\":1", json);
            Assert.Contains("\"font_italic\":0", json);
            Assert.Contains("\"font_underline\":0", json);
            Assert.Contains("\"font_strikethrough\":0", json);
        }

        [Fact]
        public void TestTextFillSerialization_SolidFill()
        {
            // Create a text component with solid fill
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
                                Text = "Red Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Solid,
                                    Color = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                                    Transparency = 0.0f
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Solid Fill JSON Output:");
            Console.WriteLine(json);
            
            // Verify text fill is present
            Assert.Contains("\"text_fill\":", json);
            Assert.Contains("\"has_fill\":1", json);
            Assert.Contains("\"fill_type\":\"solid\"", json);
            Assert.Contains("\"color\":\"RGB(255, 0, 0)\"", json);
            Assert.Contains("\"transparency\":0.00", json);
        }

        [Fact]
        public void TestTextFillSerialization_GradientFill()
        {
            // Create a text component with gradient fill
            var gradientInfo = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f
            };
            gradientInfo.Stops.Add(new GradientStop(0.0f, new ColorInfo(255, 0, 0, false)));
            gradientInfo.Stops.Add(new GradientStop(1.0f, new ColorInfo(0, 0, 255, false)));
            
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
                                Text = "Gradient Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Gradient,
                                    Gradient = gradientInfo
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Gradient Fill JSON Output:");
            Console.WriteLine(json);
            
            // Verify gradient fill is present
            Assert.Contains("\"text_fill\":", json);
            Assert.Contains("\"has_fill\":1", json);
            Assert.Contains("\"fill_type\":\"gradient\"", json);
            Assert.Contains("\"gradient_type\":\"Linear\"", json);
            Assert.Contains("\"angle\":90.0", json);
            Assert.Contains("\"stops\":[", json);
            Assert.Contains("\"position\":0.00", json);
            Assert.Contains("\"position\":1.00", json);
        }

        [Fact]
        public void TestTextFillSerialization_PatternFill()
        {
            // Create a text component with pattern fill
            var patternInfo = new PatternInfo
            {
                PatternType = "Dots",
                ForegroundColor = new ColorInfo(255, 0, 0, false),
                BackgroundColor = new ColorInfo(255, 255, 255, false)
            };
            
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
                                Text = "Pattern Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Pattern,
                                    Pattern = patternInfo
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Pattern Fill JSON Output:");
            Console.WriteLine(json);
            
            // Verify pattern fill is present
            Assert.Contains("\"text_fill\":", json);
            Assert.Contains("\"has_fill\":1", json);
            Assert.Contains("\"fill_type\":\"pattern\"", json);
            Assert.Contains("\"pattern_type\":\"Dots\"", json);
            Assert.Contains("\"foreground_color\":\"RGB(255, 0, 0)\"", json);
            Assert.Contains("\"background_color\":\"RGB(255, 255, 255)\"", json);
        }

        [Fact]
        public void TestTextFillSerialization_ThemeColor()
        {
            // Create a text component with theme color fill
            var themeColor = new ColorInfo(75, 172, 198, false)
            {
                IsThemeColor = true,
                SchemeColorName = "accent1",
                Transforms = new ColorTransforms
                {
                    LumMod = 75000,
                    LumOff = 25000
                }
            };
            
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
                                Text = "Theme Color Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Solid,
                                    Color = themeColor,
                                    Transparency = 0.0f
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Theme Color Fill JSON Output:");
            Console.WriteLine(json);
            
            // Verify theme color is preserved
            Assert.Contains("\"text_fill\":", json);
            Assert.Contains("\"has_fill\":1", json);
            Assert.Contains("\"fill_type\":\"solid\"", json);
            Assert.Contains("\"schemeColor\":\"accent1\"", json);
            Assert.Contains("\"colorTransforms\":", json);
            Assert.Contains("\"lumMod\":75000", json);
            Assert.Contains("\"lumOff\":25000", json);
        }

        [Fact]
        public void TestTextFillSerialization_NoFill()
        {
            // Create a text component with no fill
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
                                Text = "No Fill Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = false,
                                    FillType = FillType.NoFill
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("No Fill JSON Output:");
            Console.WriteLine(json);
            
            // Verify no fill is NOT serialized (backward compatibility - requirement 5.3)
            // When HasFill is false, the text_fill property should not be present
            Assert.DoesNotContain("\"text_fill\":", json);
        }

        [Fact]
        public void TestTextOutlineSerialization_WithOutline()
        {
            // Create a text component with outline
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
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextOutline = new TextOutlineInfo
                                {
                                    HasOutline = true,
                                    Width = 2.0f,
                                    Color = new ColorInfo { Red = 0, Green = 0, Blue = 255 },
                                    DashStyle = LineDashStyle.Solid,
                                    CompoundLineType = "Single",
                                    CapType = "Flat",
                                    JoinType = "Round",
                                    Transparency = 0.0f
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Text Outline JSON Output:");
            Console.WriteLine(json);
            
            // Verify text outline is present
            Assert.Contains("\"text_outline\":", json);
            Assert.Contains("\"has_outline\":1", json);
            Assert.Contains("\"width\":2.00", json);
            Assert.Contains("\"color\":\"RGB(0, 0, 255)\"", json);
            Assert.Contains("\"dash_style\":\"solid\"", json);
            Assert.Contains("\"compound_type\":\"single\"", json);
            Assert.Contains("\"cap_type\":\"flat\"", json);
            Assert.Contains("\"join_type\":\"round\"", json);
            Assert.Contains("\"transparency\":0.00", json);
        }

        [Fact]
        public void TestTextOutlineSerialization_NoOutline()
        {
            // Create a text component with no outline
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
                                Text = "No Outline Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextOutline = new TextOutlineInfo
                                {
                                    HasOutline = false
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("No Outline JSON Output:");
            Console.WriteLine(json);
            
            // Verify no outline is NOT serialized (backward compatibility - requirement 5.3)
            // When HasOutline is false, the text_outline property should not be present
            Assert.DoesNotContain("\"text_outline\":", json);
        }

        [Fact]
        public void TestTextOutlineSerialization_ThemeColor()
        {
            // Create a text component with theme color outline
            var themeColor = new ColorInfo(75, 172, 198, false)
            {
                IsThemeColor = true,
                SchemeColorName = "accent2",
                Transforms = new ColorTransforms
                {
                    LumMod = 60000,
                    LumOff = 40000
                }
            };
            
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
                                Text = "Theme Outline Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextOutline = new TextOutlineInfo
                                {
                                    HasOutline = true,
                                    Width = 1.5f,
                                    Color = themeColor,
                                    DashStyle = LineDashStyle.Dash,
                                    CompoundLineType = "Double",
                                    CapType = "Round",
                                    JoinType = "Miter",
                                    Transparency = 0.2f
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Theme Color Outline JSON Output:");
            Console.WriteLine(json);
            
            // Verify theme color outline is preserved
            Assert.Contains("\"text_outline\":", json);
            Assert.Contains("\"has_outline\":1", json);
            Assert.Contains("\"width\":1.50", json);
            Assert.Contains("\"schemeColor\":\"accent2\"", json);
            Assert.Contains("\"colorTransforms\":", json);
            Assert.Contains("\"lumMod\":60000", json);
            Assert.Contains("\"lumOff\":40000", json);
            Assert.Contains("\"dash_style\":\"dash\"", json);
            Assert.Contains("\"compound_type\":\"double\"", json);
            Assert.Contains("\"cap_type\":\"round\"", json);
            Assert.Contains("\"join_type\":\"miter\"", json);
            Assert.Contains("\"transparency\":0.20", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_OuterShadow()
        {
            // Create a text component with outer shadow
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
                                Text = "Shadow Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = new ShadowInfo
                                    {
                                        HasShadow = true,
                                        Type = ShadowType.Outer,
                                        Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                        Blur = 4.0f,
                                        Distance = 3.0f,
                                        Angle = 45.0f,
                                        Transparency = 50.0f
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Outer Shadow JSON Output:");
            Console.WriteLine(json);
            
            // Verify text effects with shadow is present
            Assert.Contains("\"text_effects\":", json);
            Assert.Contains("\"has_effects\":1", json);
            Assert.Contains("\"shadow\":", json);
            Assert.Contains("\"has_shadow\":1", json);
            Assert.Contains("\"type\":\"outer\"", json);
            Assert.Contains("\"color\":\"RGB(0, 0, 0)\"", json);
            Assert.Contains("\"blur\":4.00", json);
            Assert.Contains("\"distance\":3.00", json);
            Assert.Contains("\"angle\":45.00", json);
            Assert.Contains("\"transparency\":50.00", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_InnerShadow()
        {
            // Create a text component with inner shadow
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
                                Text = "Inner Shadow Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = new ShadowInfo
                                    {
                                        HasShadow = true,
                                        Type = ShadowType.Inner,
                                        Color = new ColorInfo { Red = 128, Green = 128, Blue = 128 },
                                        Blur = 2.0f,
                                        Distance = 1.5f,
                                        Angle = 90.0f,
                                        Transparency = 30.0f
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Inner Shadow JSON Output:");
            Console.WriteLine(json);
            
            // Verify inner shadow is present
            Assert.Contains("\"text_effects\":", json);
            Assert.Contains("\"has_effects\":1", json);
            Assert.Contains("\"shadow\":", json);
            Assert.Contains("\"has_shadow\":1", json);
            Assert.Contains("\"type\":\"inner\"", json);
            Assert.Contains("\"color\":\"RGB(128, 128, 128)\"", json);
            Assert.Contains("\"blur\":2.00", json);
            Assert.Contains("\"distance\":1.50", json);
            Assert.Contains("\"angle\":90.00", json);
            Assert.Contains("\"transparency\":30.00", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_Glow()
        {
            // Create a text component with glow effect
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
                                Text = "Glow Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasGlow = true,
                                    Glow = new GlowInfo
                                    {
                                        Radius = 5.0f,
                                        Color = new ColorInfo { Red = 255, Green = 255, Blue = 0 },
                                        Transparency = 0.0f
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Glow Effect JSON Output:");
            Console.WriteLine(json);
            
            // Verify glow effect is present
            Assert.Contains("\"text_effects\":", json);
            Assert.Contains("\"has_effects\":1", json);
            Assert.Contains("\"glow\":", json);
            Assert.Contains("\"has_glow\":1", json);
            Assert.Contains("\"radius\":5.00", json);
            Assert.Contains("\"color\":\"RGB(255, 255, 0)\"", json);
            Assert.Contains("\"transparency\":0.00", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_Reflection()
        {
            // Create a text component with reflection effect
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
                                Text = "Reflection Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasReflection = true,
                                    Reflection = new ReflectionInfo
                                    {
                                        BlurRadius = 2.5f,
                                        StartOpacity = 1.0f,
                                        StartPosition = 0.0f,
                                        EndAlpha = 0.0f,
                                        EndPosition = 1.0f,
                                        Distance = 0.0f,
                                        Direction = 0.0f,
                                        FadeDirection = 90.0f,
                                        SkewHorizontal = 0.0f,
                                        SkewVertical = 0.0f
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Reflection Effect JSON Output:");
            Console.WriteLine(json);
            
            // Verify reflection effect is present
            Assert.Contains("\"text_effects\":", json);
            Assert.Contains("\"has_effects\":1", json);
            Assert.Contains("\"reflection\":", json);
            Assert.Contains("\"has_reflection\":1", json);
            Assert.Contains("\"blur_radius\":2.50", json);
            Assert.Contains("\"start_opacity\":1.00", json);
            Assert.Contains("\"start_position\":0.00", json);
            Assert.Contains("\"end_alpha\":0.00", json);
            Assert.Contains("\"end_position\":1.00", json);
            Assert.Contains("\"distance\":0.00", json);
            Assert.Contains("\"direction\":0.00", json);
            Assert.Contains("\"fade_direction\":90.00", json);
            Assert.Contains("\"skew_horizontal\":0.00", json);
            Assert.Contains("\"skew_vertical\":0.00", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_SoftEdge()
        {
            // Create a text component with soft edge effect
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
                                Text = "Soft Edge Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasSoftEdge = true,
                                    SoftEdgeRadius = 3.5f
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Soft Edge Effect JSON Output:");
            Console.WriteLine(json);
            
            // Verify soft edge effect is present
            Assert.Contains("\"text_effects\":", json);
            Assert.Contains("\"has_effects\":1", json);
            Assert.Contains("\"soft_edge\":", json);
            Assert.Contains("\"has_soft_edge\":1", json);
            Assert.Contains("\"radius\":3.50", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_MultipleEffects()
        {
            // Create a text component with multiple effects (shadow + glow)
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
                                Text = "Multiple Effects Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = new ShadowInfo
                                    {
                                        HasShadow = true,
                                        Type = ShadowType.Outer,
                                        Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                        Blur = 4.0f,
                                        Distance = 3.0f,
                                        Angle = 45.0f,
                                        Transparency = 50.0f
                                    },
                                    HasGlow = true,
                                    Glow = new GlowInfo
                                    {
                                        Radius = 5.0f,
                                        Color = new ColorInfo { Red = 255, Green = 255, Blue = 0 },
                                        Transparency = 0.0f
                                    }
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("Multiple Effects JSON Output:");
            Console.WriteLine(json);
            
            // Verify both effects are present
            Assert.Contains("\"text_effects\":", json);
            Assert.Contains("\"has_effects\":1", json);
            
            // Verify shadow
            Assert.Contains("\"shadow\":", json);
            Assert.Contains("\"has_shadow\":1", json);
            Assert.Contains("\"type\":\"outer\"", json);
            
            // Verify glow
            Assert.Contains("\"glow\":", json);
            Assert.Contains("\"has_glow\":1", json);
            Assert.Contains("\"radius\":5.00", json);
        }

        [Fact]
        public void TestTextEffectsSerialization_NoEffects()
        {
            // Create a text component with no effects
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
                                Text = "No Effects Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = false
                                }
                            }
                        }
                    }
                }
            };
            
            // Serialize to JSON
            var json = textComponent.ToJson();
            
            // Output for manual inspection
            Console.WriteLine("No Effects JSON Output:");
            Console.WriteLine(json);
            
            // Verify no effects is NOT serialized (backward compatibility - requirement 5.3)
            // When HasEffects is false, the text_effects property should not be present
            Assert.DoesNotContain("\"text_effects\":", json);
        }
    }
}
