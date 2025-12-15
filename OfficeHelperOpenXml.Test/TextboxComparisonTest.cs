using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using OfficeHelperOpenXml.Api;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Comparison tests between original textbox.pptx and generated textbox_from_json.pptx
    /// Tests Requirement: 14.3
    /// </summary>
    public class TextboxComparisonTest
    {
        private readonly string _workspaceRoot;
        private readonly string _originalPptxPath;
        private readonly string _generatedPptxPath;

        public TextboxComparisonTest()
        {
            _workspaceRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
            _originalPptxPath = Path.Combine(_workspaceRoot, "test_ppt", "textbox.pptx");
            _generatedPptxPath = Path.Combine(_workspaceRoot, "test_ppt", "textbox_from_json.pptx");
        }

        [Fact]
        public void BothFiles_Exist()
        {
            Assert.True(File.Exists(_originalPptxPath), $"Original PPTX not found: {_originalPptxPath}");
            Assert.True(File.Exists(_generatedPptxPath), $"Generated PPTX not found: {_generatedPptxPath}");
        }

        [Fact]
        public void SlideCount_Matches()
        {
            using (var originalDoc = PresentationDocument.Open(_originalPptxPath, false))
            using (var generatedDoc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var originalSlideCount = originalDoc.PresentationPart.Presentation.SlideIdList?.Count() ?? 0;
                var generatedSlideCount = generatedDoc.PresentationPart.Presentation.SlideIdList?.Count() ?? 0;

                Assert.Equal(originalSlideCount, generatedSlideCount);
            }
        }

        [Fact]
        public void MasterSlideCount_Matches()
        {
            using (var originalDoc = PresentationDocument.Open(_originalPptxPath, false))
            using (var generatedDoc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var originalMasterCount = originalDoc.PresentationPart.Presentation.SlideMasterIdList?.Count() ?? 0;
                var generatedMasterCount = generatedDoc.PresentationPart.Presentation.SlideMasterIdList?.Count() ?? 0;

                Assert.Equal(originalMasterCount, generatedMasterCount);
            }
        }

        [Fact]
        public void SlideSize_Matches()
        {
            using (var originalDoc = PresentationDocument.Open(_originalPptxPath, false))
            using (var generatedDoc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var originalSize = originalDoc.PresentationPart.Presentation.SlideSize;
                var generatedSize = generatedDoc.PresentationPart.Presentation.SlideSize;

                Assert.NotNull(originalSize);
                Assert.NotNull(generatedSize);
                
                // Slide dimensions should match
                Assert.Equal(originalSize.Cx?.Value, generatedSize.Cx?.Value);
                Assert.Equal(originalSize.Cy?.Value, generatedSize.Cy?.Value);
            }
        }

        [Fact]
        public void PowerPointReader_CanReadBothFiles()
        {
            // Read original
            using (var originalReader = PowerPointReaderFactory.CreateReader(_originalPptxPath, out bool originalSuccess))
            {
                Assert.True(originalSuccess, "Failed to read original PPTX");
                
                var originalInfo = originalReader.PresentationInfo;
                Assert.NotNull(originalInfo);
                Assert.NotNull(originalInfo.Slides);
                
                // Read generated
                using (var generatedReader = PowerPointReaderFactory.CreateReader(_generatedPptxPath, out bool generatedSuccess))
                {
                    Assert.True(generatedSuccess, "Failed to read generated PPTX");
                    
                    var generatedInfo = generatedReader.PresentationInfo;
                    Assert.NotNull(generatedInfo);
                    Assert.NotNull(generatedInfo.Slides);
                    
                    // Compare slide counts
                    Assert.Equal(originalInfo.Slides.Count, generatedInfo.Slides.Count);
                    
                    // Compare slide dimensions
                    Assert.Equal(originalInfo.SlideWidth, generatedInfo.SlideWidth);
                    Assert.Equal(originalInfo.SlideHeight, generatedInfo.SlideHeight);
                }
            }
        }

        [Fact]
        public void ShapeCount_PerSlide_IsReasonable()
        {
            using (var originalDoc = PresentationDocument.Open(_originalPptxPath, false))
            using (var generatedDoc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var originalSlides = originalDoc.PresentationPart.Presentation.SlideIdList;
                var generatedSlides = generatedDoc.PresentationPart.Presentation.SlideIdList;

                Assert.NotNull(originalSlides);
                Assert.NotNull(generatedSlides);
                Assert.Equal(originalSlides.Count(), generatedSlides.Count());

                // Compare shape counts for each slide
                var originalSlideIds = originalSlides.Elements<DocumentFormat.OpenXml.Presentation.SlideId>().ToList();
                var generatedSlideIds = generatedSlides.Elements<DocumentFormat.OpenXml.Presentation.SlideId>().ToList();

                for (int i = 0; i < originalSlideIds.Count; i++)
                {
                    var originalSlidePart = (SlidePart)originalDoc.PresentationPart.GetPartById(originalSlideIds[i].RelationshipId);
                    var generatedSlidePart = (SlidePart)generatedDoc.PresentationPart.GetPartById(generatedSlideIds[i].RelationshipId);

                    var originalShapeTree = originalSlidePart.Slide.CommonSlideData.ShapeTree;
                    var generatedShapeTree = generatedSlidePart.Slide.CommonSlideData.ShapeTree;

                    int originalShapeCount = CountShapes(originalShapeTree);
                    int generatedShapeCount = CountShapes(generatedShapeTree);

                    // The generated slide should have the same number of shapes as the original
                    Assert.Equal(originalShapeCount, generatedShapeCount);
                }
            }
        }

        private int CountShapes(DocumentFormat.OpenXml.Presentation.ShapeTree shapeTree)
        {
            return shapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Count() +
                   shapeTree.Elements<DocumentFormat.OpenXml.Presentation.Picture>().Count() +
                   shapeTree.Elements<DocumentFormat.OpenXml.Presentation.GraphicFrame>().Count() +
                   shapeTree.Elements<DocumentFormat.OpenXml.Presentation.ConnectionShape>().Count();
        }

        [Fact]
        public void GeneratedFile_IsNotSignificantlyLarger()
        {
            var originalFileInfo = new FileInfo(_originalPptxPath);
            var generatedFileInfo = new FileInfo(_generatedPptxPath);

            // Generated file should not be more than 10x the size of original
            // (allowing for some overhead from the conversion process)
            Assert.True(generatedFileInfo.Length < originalFileInfo.Length * 10,
                $"Generated file ({generatedFileInfo.Length} bytes) is significantly larger than original ({originalFileInfo.Length} bytes)");
        }
    }
}
