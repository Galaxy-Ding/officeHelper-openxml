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
    /// Property-based tests for text style definitions
    /// Feature: pptx-phase2-implementation-fixes
    /// </summary>
    public class TextStyleDefinitionsPropertyTests
    {
        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 10: All paragraphs have endParaRPr
        /// Validates: Requirements 11.3, 11.4, 11.5
        /// 
        /// Property: For any generated paragraph, the paragraph should end with an EndParagraphRunProperties element.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property AllParagraphsHaveEndParaRPr()
        {
            return Prop.ForAll(
                GeneratePresentationWithText(),
                presentationData =>
                {
                    // Skip null data
                    if (presentationData == null)
                        return true.ToProperty().Label("Null data skipped");

                    // Create a temporary output file
                    var tempOutputPath = Path.Combine(Path.GetTempPath(), $"test_endpara_{Guid.NewGuid()}.pptx");
                    var tempJsonPath = Path.Combine(Path.GetTempPath(), $"test_endpara_{Guid.NewGuid()}.json");

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

                        // Analyze paragraphs
                        var analysis = AnalyzeParagraphs(tempOutputPath);

                        if (analysis.TotalParagraphs == 0)
                            return true.ToProperty().Label("No paragraphs found");

                        // Check that all paragraphs have endParaRPr
                        if (analysis.ParagraphsWithoutEndParaRPr > 0)
                        {
                            return false.ToProperty().Label($"FAIL: {analysis.ParagraphsWithoutEndParaRPr} of {analysis.TotalParagraphs} paragraphs missing endParaRPr");
                        }

                        return true.ToProperty().Label($"PASS: All {analysis.TotalParagraphs} paragraphs have endParaRPr");
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
        /// Feature: pptx-phase2-implementation-fixes, Property 11: Presentations have defaultTextStyle
        /// Validates: Requirements 11.1
        /// 
        /// Property: For any generated presentation, the Presentation element should contain a defaultTextStyle element.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property PresentationsHaveDefaultTextStyle()
        {
            return Prop.ForAll(
                GeneratePresentationWithText(),
                presentationData =>
                {
                    // Skip null data
                    if (presentationData == null)
                        return true.ToProperty().Label("Null data skipped");

                    // Create a temporary output file
                    var tempOutputPath = Path.Combine(Path.GetTempPath(), $"test_defaultstyle_{Guid.NewGuid()}.pptx");
                    var tempJsonPath = Path.Combine(Path.GetTempPath(), $"test_defaultstyle_{Guid.NewGuid()}.json");

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

                        // Check for defaultTextStyle in presentation.xml
                        var hasDefaultTextStyle = CheckDefaultTextStyle(tempOutputPath);

                        if (!hasDefaultTextStyle)
                        {
                            return false.ToProperty().Label("FAIL: Presentation missing defaultTextStyle element");
                        }

                        return true.ToProperty().Label("PASS: Presentation has defaultTextStyle");
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
        /// Feature: pptx-phase2-implementation-fixes, Property 12: Master slides have txStyles
        /// Validates: Requirements 11.2
        /// 
        /// Property: For any generated master slide, the SlideMaster element should contain a TextStyles element.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property MasterSlidesHaveTxStyles()
        {
            return Prop.ForAll(
                GeneratePresentationWithText(),
                presentationData =>
                {
                    // Skip null data
                    if (presentationData == null)
                        return true.ToProperty().Label("Null data skipped");

                    // Create a temporary output file
                    var tempOutputPath = Path.Combine(Path.GetTempPath(), $"test_txstyles_{Guid.NewGuid()}.pptx");
                    var tempJsonPath = Path.Combine(Path.GetTempPath(), $"test_txstyles_{Guid.NewGuid()}.json");

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

                        // Analyze master slides
                        var analysis = AnalyzeMasterSlides(tempOutputPath);

                        if (analysis.TotalMasters == 0)
                            return true.ToProperty().Label("No master slides found");

                        // Check that all master slides have txStyles
                        if (analysis.MastersWithoutTxStyles > 0)
                        {
                            return false.ToProperty().Label($"FAIL: {analysis.MastersWithoutTxStyles} of {analysis.TotalMasters} masters missing txStyles");
                        }

                        return true.ToProperty().Label($"PASS: All {analysis.TotalMasters} master slides have txStyles");
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
        /// Analyzes paragraphs in a PPTX file and counts missing endParaRPr elements
        /// </summary>
        private static ParagraphAnalysis AnalyzeParagraphs(string pptxPath)
        {
            var analysis = new ParagraphAnalysis();

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

                        // Find all paragraphs (a:p)
                        var paragraphs = doc.Descendants(ns + "p");

                        foreach (var paragraph in paragraphs)
                        {
                            analysis.TotalParagraphs++;

                            // Check for endParaRPr element
                            var endParaRPr = paragraph.Elements(ns + "endParaRPr").FirstOrDefault();
                            if (endParaRPr == null)
                            {
                                analysis.ParagraphsWithoutEndParaRPr++;
                            }
                        }
                    }
                }
            }

            return analysis;
        }

        /// <summary>
        /// Checks if presentation.xml contains defaultTextStyle element
        /// </summary>
        private static bool CheckDefaultTextStyle(string pptxPath)
        {
            using (var archive = ZipFile.OpenRead(pptxPath))
            {
                var presentationEntry = archive.GetEntry("ppt/presentation.xml");
                if (presentationEntry == null)
                    return false;

                using (var stream = presentationEntry.Open())
                {
                    var doc = XDocument.Load(stream);
                    var ns = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");

                    // Check for defaultTextStyle element
                    var defaultTextStyle = doc.Descendants(ns + "defaultTextStyle").FirstOrDefault();
                    return defaultTextStyle != null;
                }
            }
        }

        /// <summary>
        /// Analyzes master slides in a PPTX file and counts missing txStyles elements
        /// </summary>
        private static MasterSlideAnalysis AnalyzeMasterSlides(string pptxPath)
        {
            var analysis = new MasterSlideAnalysis();

            using (var archive = ZipFile.OpenRead(pptxPath))
            {
                // Find all master slide XML files
                var masterEntries = archive.Entries.Where(e => 
                    e.FullName.StartsWith("ppt/slideMasters/slideMaster") && 
                    e.FullName.EndsWith(".xml"));

                foreach (var entry in masterEntries)
                {
                    analysis.TotalMasters++;

                    using (var stream = entry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        var ns = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");

                        // Check for txStyles element
                        var txStyles = doc.Descendants(ns + "txStyles").FirstOrDefault();
                        if (txStyles == null)
                        {
                            analysis.MastersWithoutTxStyles++;
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

            // Generate presentation with content slides only
            var presentationGen = from contentSlides in Gen.NonEmptyListOf(slideGen).Select(s => s.Take(2).ToList())
                                 select new PresentationJsonData
                                 {
                                     ContentSlides = contentSlides
                                 };

            return Arb.From(presentationGen);
        }

        /// <summary>
        /// Helper class to track paragraph analysis results
        /// </summary>
        private class ParagraphAnalysis
        {
            public int TotalParagraphs { get; set; }
            public int ParagraphsWithoutEndParaRPr { get; set; }
        }

        /// <summary>
        /// Helper class to track master slide analysis results
        /// </summary>
        private class MasterSlideAnalysis
        {
            public int TotalMasters { get; set; }
            public int MastersWithoutTxStyles { get; set; }
        }
    }
}
