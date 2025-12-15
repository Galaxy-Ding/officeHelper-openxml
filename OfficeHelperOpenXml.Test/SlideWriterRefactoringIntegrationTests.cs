using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Core.Writers;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Integration tests for SlideWriter refactoring
    /// Feature: pptx-phase2-implementation-fixes
    /// Validates: Requirements 1.1, 1.2, 2.1, 2.2
    /// </summary>
    public class SlideWriterRefactoringIntegrationTests
    {
        private static readonly string TestJsonPath = Path.Combine("..", "..", "..", "..", "test_ppt", "textbox.json");
        private static readonly string TestOutputPath = Path.Combine("..", "..", "..", "..", "test_ppt", "textbox_phase2_test.pptx");

        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 1: SlideWriter receives RelationshipIdManager
        /// Validates: Requirements 1.1, 1.2, 2.1, 2.2
        /// 
        /// Test that JsonToPptxConverter creates SlideWriter with RelationshipIdManager
        /// </summary>
        [Fact]
        public void JsonToPptxConverter_CreatesSlideWriter_WithRelationshipIdManager()
        {
            // Arrange
            var converter = new JsonToPptxConverter();

            // Act - Use reflection to access private _slideWriter field
            var slideWriterField = typeof(JsonToPptxConverter).GetField("_slideWriter", 
                BindingFlags.NonPublic | BindingFlags.Instance);
            
            Assert.NotNull(slideWriterField);
            var slideWriter = slideWriterField.GetValue(converter);

            // Assert
            Assert.NotNull(slideWriter);
            Assert.IsType<SlideWriter>(slideWriter);

            // Verify SlideWriter has RelationshipIdManager
            var relationshipIdManagerField = typeof(SlideWriter).GetField("_relationshipIdManager",
                BindingFlags.NonPublic | BindingFlags.Instance);
            
            Assert.NotNull(relationshipIdManagerField);
            var relationshipIdManager = relationshipIdManagerField.GetValue(slideWriter);
            
            Assert.NotNull(relationshipIdManager);
            Assert.IsType<RelationshipIdManager>(relationshipIdManager);
        }

        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 1: SlideWriter receives RelationshipIdManager
        /// Validates: Requirements 1.1, 1.2, 2.1, 2.2
        /// 
        /// Test that generated PPTX has no GUID-based relationship IDs
        /// </summary>
        [Fact]
        public void GeneratedPptx_HasNoGuidBasedRelationshipIds()
        {
            // Arrange
            if (File.Exists(TestOutputPath))
            {
                File.Delete(TestOutputPath);
            }

            var converter = new JsonToPptxConverter();

            // Act
            bool success = converter.Convert(TestJsonPath, TestOutputPath);

            // Assert
            Assert.True(success, "Conversion should succeed");
            Assert.True(File.Exists(TestOutputPath), "Output file should exist");

            // Extract all relationship IDs from the PPTX
            var relationshipIds = ExtractAllRelationshipIds(TestOutputPath);
            
            Assert.NotEmpty(relationshipIds);

            // Check that no relationship ID contains hyphens (GUID pattern)
            var guidPattern = new Regex(@"-");
            var guidBasedIds = relationshipIds.Where(id => guidPattern.IsMatch(id)).ToList();

            Assert.Empty(guidBasedIds);

            // Also check for GUID-like patterns (8-4-4-4-12 hex digits)
            var fullGuidPattern = new Regex(@"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}");
            var fullGuidIds = relationshipIds.Where(id => fullGuidPattern.IsMatch(id)).ToList();

            Assert.Empty(fullGuidIds);
        }

        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes
        /// Validates: Requirements 1.4, 2.4
        /// 
        /// Test that all relationship IDs follow the sequential pattern rId\d+
        /// </summary>
        [Fact]
        public void GeneratedPptx_AllRelationshipIds_AreSequential()
        {
            // Arrange
            if (File.Exists(TestOutputPath))
            {
                File.Delete(TestOutputPath);
            }

            var converter = new JsonToPptxConverter();

            // Act
            bool success = converter.Convert(TestJsonPath, TestOutputPath);

            // Assert
            Assert.True(success, "Conversion should succeed");
            Assert.True(File.Exists(TestOutputPath), "Output file should exist");

            // Extract all relationship IDs from the PPTX
            var relationshipIds = ExtractAllRelationshipIds(TestOutputPath);
            
            Assert.NotEmpty(relationshipIds);

            // Check that all relationship IDs match the sequential pattern rId\d+
            var sequentialPattern = new Regex(@"^rId\d+$");
            
            foreach (var id in relationshipIds)
            {
                Assert.Matches(sequentialPattern, id);
            }
        }

        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes
        /// Validates: Requirements 1.3, 1.4
        /// 
        /// Test that image relationships use sequential IDs (not GUIDs)
        /// </summary>
        [Fact]
        public void GeneratedPptx_ImageRelationships_UseSequentialIds()
        {
            // Arrange
            if (File.Exists(TestOutputPath))
            {
                File.Delete(TestOutputPath);
            }

            var converter = new JsonToPptxConverter();

            // Act
            bool success = converter.Convert(TestJsonPath, TestOutputPath);

            // Assert
            Assert.True(success, "Conversion should succeed");
            Assert.True(File.Exists(TestOutputPath), "Output file should exist");

            // Open the PPTX and check image relationships
            using (var document = PresentationDocument.Open(TestOutputPath, false))
            {
                var slides = document.PresentationPart.SlideParts;
                
                foreach (var slidePart in slides)
                {
                    var imageParts = slidePart.ImageParts;
                    
                    foreach (var imagePart in imageParts)
                    {
                        var relationshipId = slidePart.GetIdOfPart(imagePart);
                        
                        // Verify the relationship ID is sequential (rId\d+)
                        var sequentialPattern = new Regex(@"^rId\d+$");
                        Assert.Matches(sequentialPattern, relationshipId);
                        
                        // Verify it doesn't contain GUID patterns
                        Assert.DoesNotContain("-", relationshipId);
                    }
                }
            }
        }

        /// <summary>
        /// Extracts all relationship IDs from a PPTX file, excluding root-level .rels files
        /// that are managed by the OpenXML SDK
        /// </summary>
        private static System.Collections.Generic.List<string> ExtractAllRelationshipIds(string pptxPath)
        {
            var relationshipIds = new System.Collections.Generic.List<string>();

            using (var archive = ZipFile.OpenRead(pptxPath))
            {
                // Find all .rels files, but exclude the root _rels/.rels file
                // which is managed by the OpenXML SDK and uses GUID-based IDs
                var relsEntries = archive.Entries.Where(e => 
                    e.FullName.EndsWith(".rels") && 
                    e.FullName != "_rels/.rels" &&
                    !e.FullName.StartsWith("docProps/"));

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
    }
}
