using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using Xunit;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Models.Json;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Verification tests for textbox.json conversion
    /// Tests Requirements: 14.2, 14.3, 14.4
    /// </summary>
    public class TextboxJsonVerificationTest
    {
        private readonly string _workspaceRoot;
        private readonly string _jsonPath;
        private readonly string _generatedPptxPath;
        private readonly PresentationJsonData _jsonData;

        public TextboxJsonVerificationTest()
        {
            _workspaceRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
            _jsonPath = Path.Combine(_workspaceRoot, "test_ppt", "textbox.json");
            _generatedPptxPath = Path.Combine(_workspaceRoot, "test_ppt", "textbox_from_json.pptx");
            
            // Load JSON data for comparison
            if (File.Exists(_jsonPath))
            {
                string jsonContent = File.ReadAllText(_jsonPath);
                _jsonData = JsonConvert.DeserializeObject<PresentationJsonData>(jsonContent);
            }
        }

        [Fact]
        public void GeneratedPptx_FileExists()
        {
            // Verify the PPTX file was created
            Assert.True(File.Exists(_generatedPptxPath), 
                $"Generated PPTX file not found at {_generatedPptxPath}");
        }

        [Fact]
        public void GeneratedPptx_CanBeOpened()
        {
            // Verify the PPTX can be opened by OpenXML SDK
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                Assert.NotNull(doc);
                Assert.NotNull(doc.PresentationPart);
                Assert.NotNull(doc.PresentationPart.Presentation);
            }
        }

        [Fact]
        public void GeneratedPptx_HasCorrectSlideCount()
        {
            // Verify slide count matches JSON
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var slideIdList = doc.PresentationPart.Presentation.SlideIdList;
                int actualSlideCount = slideIdList?.Count() ?? 0;
                int expectedSlideCount = _jsonData?.ContentSlides?.Count ?? 0;

                Assert.Equal(expectedSlideCount, actualSlideCount);
            }
        }

        [Fact]
        public void GeneratedPptx_HasMasterSlides()
        {
            // Verify master slides are present
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var masterIdList = doc.PresentationPart.Presentation.SlideMasterIdList;
                int actualMasterCount = masterIdList?.Count() ?? 0;
                int expectedMasterCount = _jsonData?.MasterSlides?.Count ?? 0;

                Assert.Equal(expectedMasterCount, actualMasterCount);
            }
        }

        [Fact]
        public void GeneratedPptx_SlidesHaveShapes()
        {
            // Verify content slides have shapes
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var slideIdList = doc.PresentationPart.Presentation.SlideIdList;
                
                foreach (var slideId in slideIdList.Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
                {
                    var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);
                    var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
                    
                    // Count shapes (excluding the group shape properties)
                    int shapeCount = shapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Count() +
                                   shapeTree.Elements<DocumentFormat.OpenXml.Presentation.Picture>().Count() +
                                   shapeTree.Elements<DocumentFormat.OpenXml.Presentation.GraphicFrame>().Count() +
                                   shapeTree.Elements<DocumentFormat.OpenXml.Presentation.ConnectionShape>().Count();
                    
                    // Each slide should have at least some shapes
                    Assert.True(shapeCount > 0, $"Slide {slideId.Id} has no shapes");
                }
            }
        }

        [Fact]
        public void GeneratedPptx_CanBeReadByPowerPointReader()
        {
            // Verify the PPTX can be read by our PowerPointReader
            using (var reader = PowerPointReaderFactory.CreateReader(_generatedPptxPath, out bool success))
            {
                Assert.True(success);
                Assert.NotNull(reader.PresentationInfo);
                
                var info = reader.PresentationInfo;
                Assert.NotNull(info.Slides);
                Assert.True(info.Slides.Count > 0);
            }
        }

        [Fact]
        public void GeneratedPptx_HasValidStructure()
        {
            // Verify the PPTX has a valid OpenXML structure
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                // Check presentation part
                Assert.NotNull(doc.PresentationPart);
                
                // Check presentation
                var presentation = doc.PresentationPart.Presentation;
                Assert.NotNull(presentation);
                
                // Check slide ID list
                Assert.NotNull(presentation.SlideIdList);
                
                // Check master slide list
                Assert.NotNull(presentation.SlideMasterIdList);
                
                // Check slide size
                Assert.NotNull(presentation.SlideSize);
            }
        }

        [Fact]
        public void GeneratedPptx_MasterSlidesHaveLayouts()
        {
            // Verify each master slide has at least one layout
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var masterIdList = doc.PresentationPart.Presentation.SlideMasterIdList;
                
                foreach (var masterId in masterIdList.Elements<DocumentFormat.OpenXml.Presentation.SlideMasterId>())
                {
                    var masterPart = (SlideMasterPart)doc.PresentationPart.GetPartById(masterId.RelationshipId);
                    var layoutIdList = masterPart.SlideMaster.SlideLayoutIdList;
                    
                    Assert.NotNull(layoutIdList);
                    Assert.True(layoutIdList.Count() > 0);
                }
            }
        }

        [Fact]
        public void GeneratedPptx_ContentSlidesLinkedToLayouts()
        {
            // Verify content slides are linked to layouts
            using (var doc = PresentationDocument.Open(_generatedPptxPath, false))
            {
                var slideIdList = doc.PresentationPart.Presentation.SlideIdList;
                
                foreach (var slideId in slideIdList.Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
                {
                    var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);
                    
                    // Check if slide has a layout part
                    var layoutPart = slidePart.SlideLayoutPart;
                    Assert.NotNull(layoutPart);
                }
            }
        }

        [Fact]
        public void GeneratedPptx_FileSize_IsReasonable()
        {
            // Verify file size is reasonable (not empty, not too large)
            var fileInfo = new FileInfo(_generatedPptxPath);
            
            Assert.True(fileInfo.Length > 1000);
            Assert.True(fileInfo.Length < 10 * 1024 * 1024);
        }
    }
}
