using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;
using FsCheck;
using FsCheck.Xunit;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Models.Json;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Property-based tests for relationship ID generation
    /// Feature: pptx-phase2-implementation-fixes
    /// </summary>
    public class RelationshipIdPropertyTests
    {
        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 2: All relationship IDs are sequential
        /// Validates: Requirements 1.4, 2.4
        /// 
        /// Property: For any generated PPTX file, all relationship IDs should match the pattern rId\d+.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property AllRelationshipIdsAreSequential()
        {
            return Prop.ForAll(
                GeneratePresentationJsonData(),
                presentationData =>
                {
                    // Skip null data
                    if (presentationData == null)
                        return true.ToProperty().Label("Null data skipped");

                    // Create a temporary output file
                    var tempOutputPath = Path.Combine(Path.GetTempPath(), $"test_sequential_{Guid.NewGuid()}.pptx");
                    var tempJsonPath = Path.Combine(Path.GetTempPath(), $"test_sequential_{Guid.NewGuid()}.json");

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

                        // Extract and analyze all .rels files
                        var relationshipIds = ExtractAllRelationshipIds(tempOutputPath);

                        if (relationshipIds.Count == 0)
                            return true.ToProperty().Label("No relationship IDs found");

                        // Check that all relationship IDs match the sequential pattern rId\d+
                        var sequentialPattern = new Regex(@"^rId\d+$");
                        var nonSequentialIds = relationshipIds.Where(id => !sequentialPattern.IsMatch(id)).ToList();

                        if (nonSequentialIds.Any())
                        {
                            var examples = string.Join(", ", nonSequentialIds.Take(5));
                            return false.ToProperty().Label($"FAIL: Found non-sequential IDs: {examples}");
                        }

                        return true.ToProperty().Label($"PASS: All {relationshipIds.Count} relationship IDs are sequential");
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
        /// Feature: pptx-phase2-implementation-fixes, Property 3: No GUID-based relationship IDs
        /// Validates: Requirements 1.5, 2.5
        /// 
        /// Property: For any generated PPTX file, no relationship ID should contain hyphens or GUID patterns.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property NoGuidBasedRelationshipIds()
        {
            return Prop.ForAll(
                GeneratePresentationJsonData(),
                presentationData =>
                {
                    // Skip null data
                    if (presentationData == null)
                        return true.ToProperty().Label("Null data skipped");

                    // Create a temporary output file
                    var tempOutputPath = Path.Combine(Path.GetTempPath(), $"test_noguid_{Guid.NewGuid()}.pptx");
                    var tempJsonPath = Path.Combine(Path.GetTempPath(), $"test_noguid_{Guid.NewGuid()}.json");

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

                        // Extract and analyze all .rels files
                        var relationshipIds = ExtractAllRelationshipIds(tempOutputPath);

                        if (relationshipIds.Count == 0)
                            return true.ToProperty().Label("No relationship IDs found");

                        // Check that no relationship ID contains hyphens (GUID pattern)
                        var guidPattern = new Regex(@"-");
                        var guidBasedIds = relationshipIds.Where(id => guidPattern.IsMatch(id)).ToList();

                        if (guidBasedIds.Any())
                        {
                            var examples = string.Join(", ", guidBasedIds.Take(5));
                            return false.ToProperty().Label($"FAIL: Found GUID-based IDs: {examples}");
                        }

                        // Also check for GUID-like patterns (8-4-4-4-12 hex digits)
                        var fullGuidPattern = new Regex(@"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
                        var fullGuidIds = relationshipIds.Where(id => fullGuidPattern.IsMatch(id)).ToList();

                        if (fullGuidIds.Any())
                        {
                            var examples = string.Join(", ", fullGuidIds.Take(5));
                            return false.ToProperty().Label($"FAIL: Found full GUID patterns: {examples}");
                        }

                        return true.ToProperty().Label($"PASS: All {relationshipIds.Count} relationship IDs are GUID-free");
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
        /// Extracts all relationship IDs from a PPTX file
        /// </summary>
        private static System.Collections.Generic.List<string> ExtractAllRelationshipIds(string pptxPath)
        {
            var relationshipIds = new System.Collections.Generic.List<string>();

            using (var archive = ZipFile.OpenRead(pptxPath))
            {
                // Find all .rels files
                var relsEntries = archive.Entries.Where(e => e.FullName.EndsWith(".rels"));

                foreach (var entry in relsEntries)
                {
                    using (var stream = entry.Open())
                    {
                        var doc = XDocument.Load(stream);
                        var ns = doc.Root?.Name.Namespace;

                        if (ns != null)
                        {
                            var relationships = doc.Descendants(ns + "Relationship");
                            foreach (var rel in relationships)
                            {
                                var id = rel.Attribute("Id")?.Value;
                                if (!string.IsNullOrEmpty(id))
                                {
                                    relationshipIds.Add(id);
                                }
                            }
                        }
                    }
                }
            }

            return relationshipIds;
        }

        /// <summary>
        /// Generator for PresentationJsonData with random slides and shapes
        /// </summary>
        private static Arbitrary<PresentationJsonData> GeneratePresentationJsonData()
        {
            // Generate simple shapes to test relationship ID generation
            var shapeGen = from shapeType in Gen.Elements("textbox", "autoshape")
                          from x in Gen.Choose(0, 20).Select(v => (float)v)
                          from y in Gen.Choose(0, 15).Select(v => (float)v)
                          from width in Gen.Choose(5, 15).Select(v => (float)v)
                          from height in Gen.Choose(3, 10).Select(v => (float)v)
                          select CreateShape(shapeType, x, y, width, height);

            var slideGen = from shapes in Gen.ListOf(shapeGen).Select(s => s.Take(3).ToList())
                          from pageNum in Gen.Choose(1, 10)
                          select new SlideJsonData
                          {
                              PageNumber = pageNum,
                              Shapes = shapes
                          };

            var presentationGen = from slides in Gen.NonEmptyListOf(slideGen).Select(s => s.Take(2).ToList())
                                 select new PresentationJsonData
                                 {
                                     ContentSlides = slides
                                 };

            return Arb.From(presentationGen);
        }

        /// <summary>
        /// Helper method to create a shape
        /// </summary>
        private static ShapeJsonData CreateShape(string shapeType, float x, float y, float width, float height)
        {
            var shape = new ShapeJsonData
            {
                Type = shapeType,
                Name = $"Shape_{shapeType}",
                Box = $"{x},{y},{width},{height}",
                HasText = 1,
                Text = new System.Collections.Generic.List<TextRunJsonData>
                {
                    new TextRunJsonData
                    {
                        Content = "Test",
                        Font = "Arial",
                        FontSize = 12.0f
                    }
                }
            };

            return shape;
        }
    }
}
