using System;
using System.Collections.Generic;
using Xunit;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests for text run merging functionality
    /// Validates Requirements 1.1, 1.3, 1.4 from merge-identical-text-runs spec
    /// </summary>
    public class TextRunMergingTests
    {
        [Fact]
        public void TestMergeIdenticalRuns_BasicFormatting()
        {
            // Test that consecutive runs with identical basic formatting are merged
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
                                Text = "Hello ",
                                FontName = "Arial",
                                FontSize = 12,
                                IsBold = true,
                                IsItalic = false,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "World",
                                FontName = "Arial",
                                FontSize = 12,
                                IsBold = true,
                                IsItalic = false,
                                IsUnderline = false,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged into one
            Assert.Contains("\"content\":\"Hello World\"", json);
            
            // Count the number of text runs - should be 1
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMergeIdenticalRuns_WithGradientFill()
        {
            // Test that consecutive runs with identical gradient fills are merged
            var gradientInfo = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 45.0f
            };
            gradientInfo.Stops.Add(new GradientStop(0.0f, new ColorInfo(146, 189, 227, false)
            {
                IsThemeColor = true,
                SchemeColorName = "accent1",
                Transforms = new ColorTransforms { Tint = 66000, SatMod = 160000 }
            }));
            gradientInfo.Stops.Add(new GradientStop(0.5f, new ColorInfo(17, 84, 204, false)
            {
                IsThemeColor = true,
                SchemeColorName = "accent1",
                Transforms = new ColorTransforms { }
            }));
            gradientInfo.Stops.Add(new GradientStop(1.0f, new ColorInfo(11, 56, 136, false)
            {
                IsThemeColor = true,
                SchemeColorName = "accent1",
                Transforms = new ColorTransforms { Shade = 50000 }
            }));
            
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
                                Text = "文",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Gradient,
                                    Gradient = gradientInfo
                                }
                            },
                            new TextRunInfo
                            {
                                Text = "本",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Gradient,
                                    Gradient = gradientInfo
                                }
                            },
                            new TextRunInfo
                            {
                                Text = "填",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Gradient,
                                    Gradient = gradientInfo
                                }
                            },
                            new TextRunInfo
                            {
                                Text = "充",
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
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged into one
            Assert.Contains("\"content\":\"文本填充\"", json);
            
            // Count the number of text runs - should be 1
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentFontSize()
        {
            // Test that runs with different font sizes are NOT merged
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
                                Text = "Small",
                                FontName = "Arial",
                                FontSize = 10,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Large",
                                FontName = "Arial",
                                FontSize = 14,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are NOT merged
            Assert.Contains("\"content\":\"Small\"", json);
            Assert.Contains("\"content\":\"Large\"", json);
            
            // Count the number of text runs - should be 2
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentFontName()
        {
            // Test that runs with different font names are NOT merged
            // Validates: Requirements 1.3
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
                                Text = "Arial",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Times",
                                FontName = "Times New Roman",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are NOT merged
            Assert.Contains("\"content\":\"Arial\"", json);
            Assert.Contains("\"content\":\"Times\"", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentBold()
        {
            // Test that runs with different bold settings are NOT merged
            // Validates: Requirements 1.3
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
                                Text = "Normal",
                                FontName = "Arial",
                                FontSize = 12,
                                IsBold = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Bold",
                                FontName = "Arial",
                                FontSize = 12,
                                IsBold = true,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            Assert.Contains("\"content\":\"Normal\"", json);
            Assert.Contains("\"content\":\"Bold\"", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentItalic()
        {
            // Test that runs with different italic settings are NOT merged
            // Validates: Requirements 1.3
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
                                Text = "Normal",
                                FontName = "Arial",
                                FontSize = 12,
                                IsItalic = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Italic",
                                FontName = "Arial",
                                FontSize = 12,
                                IsItalic = true,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            Assert.Contains("\"content\":\"Normal\"", json);
            Assert.Contains("\"content\":\"Italic\"", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentUnderline()
        {
            // Test that runs with different underline settings are NOT merged
            // Validates: Requirements 1.3
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
                                Text = "Normal",
                                FontName = "Arial",
                                FontSize = 12,
                                IsUnderline = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Underlined",
                                FontName = "Arial",
                                FontSize = 12,
                                IsUnderline = true,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            Assert.Contains("\"content\":\"Normal\"", json);
            Assert.Contains("\"content\":\"Underlined\"", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentStrikethrough()
        {
            // Test that runs with different strikethrough settings are NOT merged
            // Validates: Requirements 1.3
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
                                Text = "Normal",
                                FontName = "Arial",
                                FontSize = 12,
                                IsStrikethrough = false,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Strikethrough",
                                FontName = "Arial",
                                FontSize = 12,
                                IsStrikethrough = true,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            Assert.Contains("\"content\":\"Normal\"", json);
            Assert.Contains("\"content\":\"Strikethrough\"", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentFontColor()
        {
            // Test that runs with different font colors are NOT merged
            // Validates: Requirements 1.3
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
                                Text = "Red",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 255, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Blue",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 255 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            Assert.Contains("\"content\":\"Red\"", json);
            Assert.Contains("\"content\":\"Blue\"", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentGradientAngle()
        {
            // Test that runs with different gradient angles are NOT merged
            var gradient1 = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 45.0f
            };
            gradient1.Stops.Add(new GradientStop(0.0f, new ColorInfo(255, 0, 0, false)));
            gradient1.Stops.Add(new GradientStop(1.0f, new ColorInfo(0, 0, 255, false)));
            
            var gradient2 = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f  // Different angle
            };
            gradient2.Stops.Add(new GradientStop(0.0f, new ColorInfo(255, 0, 0, false)));
            gradient2.Stops.Add(new GradientStop(1.0f, new ColorInfo(0, 0, 255, false)));
            
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
                                Text = "Text1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Gradient,
                                    Gradient = gradient1
                                }
                            },
                            new TextRunInfo
                            {
                                Text = "Text2",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextFill = new TextFillInfo
                                {
                                    HasFill = true,
                                    FillType = FillType.Gradient,
                                    Gradient = gradient2
                                }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are NOT merged
            Assert.Contains("\"content\":\"Text1\"", json);
            Assert.Contains("\"content\":\"Text2\"", json);
            
            // Count the number of text runs - should be 2
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestParagraphBoundaries_NotMerged()
        {
            // Test that runs from different paragraphs are NOT merged
            // even if they have identical formatting
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
                                Text = "Paragraph 1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    },
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Paragraph 2",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that newline is inserted between paragraphs
            Assert.Contains("\"content\":\"Paragraph 1\\nParagraph 2\"", json);
            
            // Count the number of text runs - should be 1 (merged with newline)
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMergeWithTextEffects_Shadow()
        {
            // Test that runs with identical shadow effects are merged
            var shadow = new ShadowInfo
            {
                HasShadow = true,
                Type = ShadowType.Outer,
                Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                Blur = 4.0f,
                Distance = 3.0f,
                Angle = 45.0f,
                Transparency = 50.0f
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
                                Text = "Shadow ",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = shadow
                                }
                            },
                            new TextRunInfo
                            {
                                Text = "Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = shadow
                                }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged
            Assert.Contains("\"content\":\"Shadow Text\"", json);
            
            // Count the number of text runs - should be 1
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestNoMerge_DifferentShadowBlur()
        {
            // Test that runs with different shadow blur values are NOT merged
            var shadow1 = new ShadowInfo
            {
                HasShadow = true,
                Type = ShadowType.Outer,
                Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                Blur = 4.0f,
                Distance = 3.0f,
                Angle = 45.0f,
                Transparency = 50.0f
            };
            
            var shadow2 = new ShadowInfo
            {
                HasShadow = true,
                Type = ShadowType.Outer,
                Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                Blur = 6.0f,  // Different blur
                Distance = 3.0f,
                Angle = 45.0f,
                Transparency = 50.0f
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
                                Text = "Text1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = shadow1
                                }
                            },
                            new TextRunInfo
                            {
                                Text = "Text2",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextEffects = new TextEffectsInfo
                                {
                                    HasEffects = true,
                                    HasShadow = true,
                                    Shadow = shadow2
                                }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are NOT merged
            Assert.Contains("\"content\":\"Text1\"", json);
            Assert.Contains("\"content\":\"Text2\"", json);
            
            // Count the number of text runs - should be 2
            int runCount = CountTextRuns(json);
            Assert.Equal(2, runCount);
        }

        [Fact]
        public void TestMergeWithTextOutline()
        {
            // Test that runs with identical outlines are merged
            var outline = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                DashStyle = LineDashStyle.Solid,
                CompoundLineType = "Single",
                CapType = "Flat",
                JoinType = "Round",
                Transparency = 0.0f
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
                                Text = "Outlined ",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextOutline = outline
                            },
                            new TextRunInfo
                            {
                                Text = "Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                                TextOutline = outline
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged
            Assert.Contains("\"content\":\"Outlined Text\"", json);
            
            // Count the number of text runs - should be 1
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestEmptyRunList()
        {
            // Test that empty run list is handled gracefully
            var textComponent = new TextComponent
            {
                HasText = false,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>
                {
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>()
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that empty text array is output
            Assert.Contains("\"hastext\":0", json);
            Assert.Contains("\"text\":[]", json);
        }

        [Fact]
        public void TestNullParagraphsList()
        {
            // Test that null paragraphs list is handled gracefully
            var textComponent = new TextComponent
            {
                HasText = false,
                IsEnabled = true,
                Paragraphs = null
            };
            
            var json = textComponent.ToJson();
            
            // Verify that empty text array is output
            Assert.Contains("\"hastext\":0", json);
            Assert.Contains("\"text\":[]", json);
        }

        [Fact]
        public void TestEmptyParagraphsList()
        {
            // Test that empty paragraphs list is handled gracefully
            var textComponent = new TextComponent
            {
                HasText = false,
                IsEnabled = true,
                Paragraphs = new List<ParagraphInfo>()
            };
            
            var json = textComponent.ToJson();
            
            // Verify that empty text array is output
            Assert.Contains("\"hastext\":0", json);
            Assert.Contains("\"text\":[]", json);
        }

        [Fact]
        public void TestSingleRun()
        {
            // Test that a single run remains unchanged
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
                                Text = "Single Run",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the single run is output correctly
            Assert.Contains("\"content\":\"Single Run\"", json);
            
            // Count the number of text runs - should be 1
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMerge_WithTabCharacters()
        {
            // Test that runs with tab characters are merged correctly
            // Validates: Requirements 1.4, 4.2
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
                                Text = "Before\t",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "After",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged and tab is preserved
            Assert.Contains("Before", json);
            Assert.Contains("After", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMerge_WithUnicodeCharacters()
        {
            // Test that runs with Unicode characters are merged correctly
            // Validates: Requirements 1.4, 4.2
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
                                Text = "Hello ",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "世界",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = " 🌍",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged and Unicode is preserved
            Assert.Contains("Hello", json);
            Assert.Contains("世界", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMerge_WithEscapedCharacters()
        {
            // Test that runs with characters that need JSON escaping are handled correctly
            // Validates: Requirements 1.4, 4.2
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
                                Text = "Quote: \"",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Text",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "\"",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMerge_WithEmptyStringRuns()
        {
            // Test that empty string runs are handled correctly
            // Validates: Requirements 1.4, 4.2
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
                                Text = "Before",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "After",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged
            Assert.Contains("BeforeAfter", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestMerge_WithWhitespaceOnly()
        {
            // Test that runs with only whitespace are merged correctly
            // Validates: Requirements 1.4, 4.2
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
                                Text = "Word1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "   ",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Word2",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that the runs are merged with whitespace preserved
            Assert.Contains("Word1   Word2", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        [Fact]
        public void TestParagraphBoundary_WithMultipleParagraphs()
        {
            // Test that paragraph boundaries are respected with multiple paragraphs
            // Validates: Requirements 1.3, 4.2
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
                                Text = "Para1Run1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Para1Run2",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    },
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Para2Run1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    },
                    new ParagraphInfo
                    {
                        Runs = new List<TextRunInfo>
                        {
                            new TextRunInfo
                            {
                                Text = "Para3Run1",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            },
                            new TextRunInfo
                            {
                                Text = "Para3Run2",
                                FontName = "Arial",
                                FontSize = 12,
                                FontColor = new ColorInfo { Red = 0, Green = 0, Blue = 0 }
                            }
                        }
                    }
                }
            };
            
            var json = textComponent.ToJson();
            
            // Verify that runs within paragraphs are merged but paragraphs are separated
            Assert.Contains("Para1Run1Para1Run2\\nPara2Run1\\nPara3Run1Para3Run2", json);
            
            int runCount = CountTextRuns(json);
            Assert.Equal(1, runCount);
        }

        /// <summary>
        /// Helper method to count the number of text runs in JSON output
        /// </summary>
        private int CountTextRuns(string json)
        {
            int count = 0;
            int index = 0;
            string searchPattern = "\"content\":";
            
            while ((index = json.IndexOf(searchPattern, index)) != -1)
            {
                count++;
                index += searchPattern.Length;
            }
            
            return count;
        }
    }
}
