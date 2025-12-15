using System;
using System.Collections.Generic;
using OfficeHelperOpenXml.Core.Writers;
using OfficeHelperOpenXml.Models.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Integration tests for SlideWriter JSON conversion methods
    /// Validates that JSON data can be converted to OpenXML shapes
    /// </summary>
    public class SlideWriterJsonMethodsTest
    {
        private static SlideWriter CreateSlideWriter()
        {
            return new SlideWriter(new OfficeHelperOpenXml.Core.Converters.RelationshipIdManager());
        }

        public static void RunTests()
        {
            Console.WriteLine("=== SlideWriter JSON Methods Tests ===\n");

            TestCreateShapeFromJson();
            TestApplyFillFromJson();
            TestApplyLineFromJson();
            TestApplyShadowFromJson();
            TestCreateParagraphFromJson();
            TestCreateRunFromJson();
            TestApplyTextEffects();

            Console.WriteLine("\n=== All SlideWriter JSON Tests Completed ===");
        }

        private static void TestCreateShapeFromJson()
        {
            Console.WriteLine("Test: CreateShapeFromJson");
            try
            {
                var slideWriter = CreateSlideWriter();
                var shapeData = new ShapeJsonData
                {
                    Type = "textbox",
                    Name = "Test TextBox",
                    Box = "1.0,2.0,5.0,3.0",
                    Rotation = 45.0f,
                    HasText = 1,
                    Text = new List<TextRunJsonData>
                    {
                        new TextRunJsonData
                        {
                            Content = "Test Text",
                            Font = "Arial",
                            FontSize = 12,
                            FontBold = 1
                        }
                    },
                    Fill = new FillJsonData
                    {
                        Color = "RGB(255, 0, 0)",
                        Opacity = 1.0f
                    },
                    Line = new LineJsonData
                    {
                        HasOutline = 1,
                        Color = "RGB(0, 0, 255)",
                        Width = 2.0f
                    }
                };

                var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null);

                if (shape != null && shape is P.Shape pShape)
                {
                    Console.WriteLine("✓ Shape created successfully");
                    Console.WriteLine($"  - Shape type: {pShape.GetType().Name}");
                    Console.WriteLine($"  - Has text body: {pShape.TextBody != null}");
                }
                else
                {
                    Console.WriteLine("✗ Failed to create shape");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }

        private static void TestApplyFillFromJson()
        {
            Console.WriteLine("Test: ApplyFillFromJson");
            try
            {
                var slideWriter = CreateSlideWriter();
                var shapeProps = new P.ShapeProperties();
                var fillData = new FillJsonData
                {
                    Color = "RGB(128, 128, 128)",
                    Opacity = 0.5f
                };

                slideWriter.ApplyFillFromJson(shapeProps, fillData);

                var solidFill = shapeProps.GetFirstChild<A.SolidFill>();
                if (solidFill != null)
                {
                    Console.WriteLine("✓ Fill applied successfully");
                }
                else
                {
                    Console.WriteLine("✗ Fill not applied");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }

        private static void TestApplyLineFromJson()
        {
            Console.WriteLine("Test: ApplyLineFromJson");
            try
            {
                var slideWriter = CreateSlideWriter();
                var shapeProps = new P.ShapeProperties();
                var lineData = new LineJsonData
                {
                    HasOutline = 1,
                    Color = "RGB(0, 255, 0)",
                    Width = 3.0f
                };

                slideWriter.ApplyLineFromJson(shapeProps, lineData);

                var outline = shapeProps.GetFirstChild<A.Outline>();
                if (outline != null)
                {
                    Console.WriteLine("✓ Line applied successfully");
                    Console.WriteLine($"  - Width: {outline.Width}");
                }
                else
                {
                    Console.WriteLine("✗ Line not applied");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }

        private static void TestApplyShadowFromJson()
        {
            Console.WriteLine("Test: ApplyShadowFromJson");
            try
            {
                var slideWriter = CreateSlideWriter();
                var shapeProps = new P.ShapeProperties();
                var shadowData = new ShadowJsonData
                {
                    HasShadow = 1,
                    Color = "RGB(0, 0, 0)",
                    Blur = 5.0f,
                    OffsetX = 2.0f,
                    OffsetY = 2.0f,
                    Transparency = 0.5f
                };

                slideWriter.ApplyShadowFromJson(shapeProps, shadowData);

                var effectStyle = shapeProps.GetFirstChild<A.EffectStyle>();
                if (effectStyle != null)
                {
                    Console.WriteLine("✓ Shadow applied successfully");
                }
                else
                {
                    Console.WriteLine("✗ Shadow not applied");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }

        private static void TestCreateParagraphFromJson()
        {
            Console.WriteLine("Test: CreateParagraphFromJson");
            try
            {
                var slideWriter = CreateSlideWriter();
                var textRuns = new List<TextRunJsonData>
                {
                    new TextRunJsonData
                    {
                        Content = "First run ",
                        Font = "Arial",
                        FontSize = 12,
                        FontBold = 1
                    },
                    new TextRunJsonData
                    {
                        Content = "Second run",
                        Font = "Times New Roman",
                        FontSize = 14,
                        FontItalic = 1
                    }
                };

                var paragraph = slideWriter.CreateParagraphFromJson(textRuns);

                if (paragraph != null)
                {
                    var runs = paragraph.Elements<A.Run>();
                    int runCount = 0;
                    foreach (var run in runs) runCount++;

                    Console.WriteLine("✓ Paragraph created successfully");
                    Console.WriteLine($"  - Number of runs: {runCount}");
                }
                else
                {
                    Console.WriteLine("✗ Paragraph not created");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }

        private static void TestCreateRunFromJson()
        {
            Console.WriteLine("Test: CreateRunFromJson");
            try
            {
                var slideWriter = CreateSlideWriter();
                var runData = new TextRunJsonData
                {
                    Content = "Test Run",
                    Font = "Calibri",
                    FontSize = 16,
                    FontColor = "RGB(255, 0, 0)",
                    FontBold = 1,
                    FontItalic = 1,
                    FontUnderline = 1,
                    FontStrikethrough = 0
                };

                var run = slideWriter.CreateRunFromJson(runData);

                if (run != null && run.Text != null)
                {
                    Console.WriteLine("✓ Run created successfully");
                    Console.WriteLine($"  - Text: {run.Text.Text}");
                    Console.WriteLine($"  - Bold: {run.RunProperties?.Bold}");
                    Console.WriteLine($"  - Italic: {run.RunProperties?.Italic}");
                }
                else
                {
                    Console.WriteLine("✗ Run not created");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }

        private static void TestApplyTextEffects()
        {
            Console.WriteLine("Test: ApplyTextEffects");
            try
            {
                var slideWriter = CreateSlideWriter();
                var runProps = new A.RunProperties();
                var effects = new TextEffectsJsonData
                {
                    HasEffects = 1,
                    Shadow = new TextShadowJsonData
                    {
                        HasShadow = 1,
                        Color = "RGB(0, 0, 0)",
                        Blur = 3.0f,
                        Distance = 2.0f,
                        Angle = 45.0f,
                        Transparency = 0.5f
                    },
                    Glow = new TextGlowJsonData
                    {
                        HasGlow = 1
                    }
                };

                slideWriter.ApplyTextEffects(runProps, effects);

                var effectStyle = runProps.GetFirstChild<A.EffectStyle>();
                if (effectStyle != null)
                {
                    Console.WriteLine("✓ Text effects applied successfully");
                }
                else
                {
                    Console.WriteLine("✗ Text effects not applied");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Exception: {ex.Message}");
            }
            Console.WriteLine();
        }
    }
}
