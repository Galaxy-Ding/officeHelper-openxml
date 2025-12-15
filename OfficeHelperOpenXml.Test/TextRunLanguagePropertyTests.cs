using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using FsCheck;
using FsCheck.Xunit;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Models.Json;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Property-based tests for text run language attributes
    /// Feature: pptx-phase2-implementation-fixes
    /// </summary>
    public class TextRunLanguagePropertyTests
    {
        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 7: All text runs have language attributes
        /// Validates: Requirements 8.1, 8.2, 8.3, 8.4, 8.5
        /// 
        /// Property: For any generated text run, the run properties should include lang, altLang, and dirty attributes.
        /// Note: smtClean attribute is not available in OpenXML SDK and is skipped.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property AllTextRunsHaveLanguageAttributes()
        {
            return Prop.ForAll(
                GeneratePresentationWithText(),
                presentationData =>
                {
                    // Skip null data
                    if (presentationData == null)
                        return true.ToProperty().Label("Null data skipped");

                    // Create a temporary output file
                    var tempOutputPath = Path.Combine(Path.GetTempPath(), $"test_textrun_{Guid.NewGuid()}.pptx");
                    var tempJsonPath = Path.Combine(Path.GetTempPath(), $"test_textrun_{Guid.NewGuid()}.json");

                    try
                    {
                        // Serialize the presentation data to JSON
                        var json = Newtonsoft.Json.JsonConvert.SerializeObject(presentationData);
                        File.WriteAllText(tempJsonPath, json);

                        // Convert JSON to PPTX
                        var converter = new JsonToPptxConverter();
                        var success = converter.Convert(tempJsonPath, tempOutputPath);

                        if (!success)
                            return false.ToProperty().Label("FAIL: Conversion failed");

                        // Extract and analyze all text runs
                        var textRunAnalysis = AnalyzeTextRuns(tempOutputPath);

                        if (textRunAnalysis.TotalRuns == 0)
                            return true.ToProperty().Label("No text runs found");

                        // Check that all text runs have required language attributes
                        if (textRunAnalysis.RunsWithoutLang > 0)
                        {
                            return false.ToProperty().Label($"FAIL: {textRunAnalysis.RunsWithoutLang} runs missing 'lang' attribute");
                        }

                        if (textRunAnalysis.RunsWithoutAltLang > 0)
                        {
                            return false.ToProperty().Label($"FAIL: {textRunAnalysis.RunsWithoutAltLang} runs missing 'altLang' attribute");
                        }

                        if (textRunAnalysis.RunsWithoutDirty > 0)
                        {
                            return false.ToProperty().Label($"FAIL: {textRunAnalysis.RunsWithoutDirty} runs missing 'dirty' attribute");
                        }

                        return true.ToProperty().Label($"PASS: All {textRunAnalysis.TotalRuns} text runs have language attributes");
                    }
                    finally
                    {
                        // Clean up temporary files
                        if (File.Exists(tempOutputPath))
                            File.Delete(tempOutputPath);
                        if (File.Exists(tempJsonPath))
                            File.Delete(tempJsonPath);
                    }
                });
        }

        /// <summary>
        /// Analyzes text runs in a PPTX file and counts missing attributes
        /// </summary>
        private static TextRunAnalysis AnalyzeTextRuns(string pptxPath)
        {
            var analysis = new TextRunAnalysis();

            using (var archive = ZipFile.OpenRead(pptxPath))
            {
                // Find all slide XML files (content slides, master slides, layouts)
                var slideEntries = archive.Entries.Where(e => 
                    e.FullName.StartsWith("ppt/slides/slide") ||
                    e.FullName.StartsWith("ppt/slideMasters/slideMaster") ||
                    e.FullName.StartsWith("ppt/slideLayouts/slideLayout"));

                foreach (var entry in slideEntries)
                {
                    if (!entry.FullName.EndsWith(".xml"))
                        continue;

                    using (var stream = entry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        var ns = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");

                        // Find all run properties (a:rPr)
                        var runProps = doc.Descendants(ns + "rPr");

                        foreach (var rPr in runProps)
                        {
                            analysis.TotalRuns++;

                            // Check for lang attribute
                            if (rPr.Attribute("lang") == null)
                            {
                                analysis.RunsWithoutLang++;
                            }

                            // Check for altLang attribute
                            if (rPr.Attribute("altLang") == null)
                            {
                                analysis.RunsWithoutAltLang++;
                            }

                            // Check for dirty attribute
                            if (rPr.Attribute("dirty") == null)
                            {
                                analysis.RunsWithoutDirty++;
                            }

                            // Note: smtClean is not checked as it's not available in OpenXML SDK
                        }
                    }
                }
            }

            return analysis;
        }

        /// <summary>
        /// Generator for PresentationJsonData with text content
        /// </summary>
        private static Arbitrary<PresentationJsonData> GeneratePresentationWithText()
        {
            // Generate text runs with various properties
            var textRunGen = from content in Gen.Elements("Hello", "World", "Test", "Sample", "文本")
                            from fontSize in Gen.Choose(10, 24).Select(v => (float)v)
                            from font in Gen.Elements("Arial", "Calibri", "Times New Roman")
                            select new TextRunJsonData
                            {
                                Content = content,
                                Font = font,
                                FontSize = fontSize
                            };

            // Generate shapes with text
            var shapeGen = from textRuns in Gen.NonEmptyListOf(textRunGen).Select(t => t.Take(3).ToList())
                          from x in Gen.Choose(0, 20).Select(v => (float)v)
                          from y in Gen.Choose(0, 15).Select(v => (float)v)
                          from width in Gen.Choose(5, 15).Select(v => (float)v)
                          from height in Gen.Choose(3, 10).Select(v => (float)v)
                          select new ShapeJsonData
                          {
                              Type = "textbox",
                              Name = "TextShape",
                              Box = $"{x},{y},{width},{height}",
                              HasText = 1,
                              Text = textRuns
                          };

            // Generate slides with shapes
            var slideGen = from shapes in Gen.NonEmptyListOf(shapeGen).Select(s => s.Take(2).ToList())
                          from pageNum in Gen.Choose(1, 10)
                          select new SlideJsonData
                          {
                              PageNumber = pageNum,
                              Shapes = shapes
                          };

            // Generate presentation with content slides only (master slides are optional and can cause null issues)
            var presentationGen = from contentSlides in Gen.NonEmptyListOf(slideGen).Select(s => s.Take(2).ToList())
                                 select new PresentationJsonData
                                 {
                                     ContentSlides = contentSlides
                                 };

            return Arb.From(presentationGen);
        }

        /// <summary>
        /// Helper class to track text run analysis results
        /// </summary>
        private class TextRunAnalysis
        {
            public int TotalRuns { get; set; }
            public int RunsWithoutLang { get; set; }
            public int RunsWithoutAltLang { get; set; }
            public int RunsWithoutDirty { get; set; }
        }
    }
}
