using System;
using System.IO;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Core.Converters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for layout relationship generation
    /// **Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5**
    /// **Property 5: Layout relationship files exist**
    /// </summary>
    public class LayoutRelationshipGenerationTests : IDisposable
    {
        private readonly string _testOutputPath;

        public LayoutRelationshipGenerationTests()
        {
            _testOutputPath = Path.Combine(Path.GetTempPath(), $"test_layout_{Guid.NewGuid()}.pptx");
        }

        public void Dispose()
        {
            if (File.Exists(_testOutputPath))
            {
                try
                {
                    File.Delete(_testOutputPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_CreatesRelationshipFile()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                masterPart.SlideMaster = CreateSlideMaster();

                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rId2");
                layoutPart.SlideLayout = CreateSlideLayout();

                var generator = new LayoutRelationshipGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                generator.GenerateLayoutRelationships(layoutPart, masterPart, relationshipIdManager);

                // Assert - Check that relationship exists
                var relationships = layoutPart.Parts;
                Assert.NotEmpty(relationships);
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_ContainsRelationshipToMaster()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                masterPart.SlideMaster = CreateSlideMaster();

                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rId2");
                layoutPart.SlideLayout = CreateSlideLayout();

                var generator = new LayoutRelationshipGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                generator.GenerateLayoutRelationships(layoutPart, masterPart, relationshipIdManager);

                // Assert - Check that relationship to master exists
                var masterRelationship = layoutPart.Parts.FirstOrDefault(p => p.OpenXmlPart == masterPart);
                Assert.NotNull(masterRelationship.OpenXmlPart);
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_UsesSequentialId()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                masterPart.SlideMaster = CreateSlideMaster();

                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rId2");
                layoutPart.SlideLayout = CreateSlideLayout();

                var generator = new LayoutRelationshipGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                generator.GenerateLayoutRelationships(layoutPart, masterPart, relationshipIdManager);

                // Assert - Check that relationship ID is sequential (rId1 after reset)
                var relationshipId = layoutPart.GetIdOfPart(masterPart);
                Assert.NotNull(relationshipId);
                Assert.Matches(@"^rId\d+$", relationshipId);
                Assert.Equal("rId1", relationshipId); // Should be rId1 after reset
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_WithNullLayoutPart_ThrowsArgumentNullException()
        {
            // Arrange
            var generator = new LayoutRelationshipGenerator();
            var relationshipIdManager = new RelationshipIdManager();

            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");

                // Act & Assert
                var exception = Assert.Throws<ArgumentNullException>(() =>
                    generator.GenerateLayoutRelationships(null!, masterPart, relationshipIdManager));
                Assert.Equal("layoutPart", exception.ParamName);
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_WithNullMasterPart_ThrowsArgumentNullException()
        {
            // Arrange
            var generator = new LayoutRelationshipGenerator();
            var relationshipIdManager = new RelationshipIdManager();

            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rId2");

                // Act & Assert
                var exception = Assert.Throws<ArgumentNullException>(() =>
                    generator.GenerateLayoutRelationships(layoutPart, null!, relationshipIdManager));
                Assert.Equal("masterPart", exception.ParamName);
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_WithNullRelationshipIdManager_ThrowsArgumentNullException()
        {
            // Arrange
            var generator = new LayoutRelationshipGenerator();

            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rId2");

                // Act & Assert
                var exception = Assert.Throws<ArgumentNullException>(() =>
                    generator.GenerateLayoutRelationships(layoutPart, masterPart, null!));
                Assert.Equal("relationshipIdManager", exception.ParamName);
            }
        }

        [Fact]
        public void GenerateLayoutRelationships_FileExistsInPackage()
        {
            // Arrange & Act
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                var presentationPart = document.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                var masterPart = presentationPart.AddNewPart<SlideMasterPart>("rId1");
                masterPart.SlideMaster = CreateSlideMaster();

                var layoutPart = masterPart.AddNewPart<SlideLayoutPart>("rId2");
                layoutPart.SlideLayout = CreateSlideLayout();

                var generator = new LayoutRelationshipGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                generator.GenerateLayoutRelationships(layoutPart, masterPart, relationshipIdManager);
                document.Save();
            }

            // Assert - Reopen and verify relationship exists
            using (var document = PresentationDocument.Open(_testOutputPath, false))
            {
                var presentationPart = document.PresentationPart;
                Assert.NotNull(presentationPart);

                var masterPart = presentationPart.SlideMasterParts.FirstOrDefault();
                Assert.NotNull(masterPart);

                var layoutPart = masterPart.SlideLayoutParts.FirstOrDefault();
                Assert.NotNull(layoutPart);

                // Verify relationship exists
                var relationships = layoutPart.Parts;
                Assert.NotEmpty(relationships);
            }
        }

        private SlideMaster CreateSlideMaster()
        {
            return new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = 0, Cy = 0 },
                        new A.ChildOffset { X = 0, Y = 0 },
                        new A.ChildExtents { Cx = 12192000, Cy = 6858000 }
                    ))
                )),
                new P.ColorMap
                {
                    Background1 = A.ColorSchemeIndexValues.Light1,
                    Text1 = A.ColorSchemeIndexValues.Dark1,
                    Background2 = A.ColorSchemeIndexValues.Light2,
                    Text2 = A.ColorSchemeIndexValues.Dark2,
                    Accent1 = A.ColorSchemeIndexValues.Accent1,
                    Accent2 = A.ColorSchemeIndexValues.Accent2,
                    Accent3 = A.ColorSchemeIndexValues.Accent3,
                    Accent4 = A.ColorSchemeIndexValues.Accent4,
                    Accent5 = A.ColorSchemeIndexValues.Accent5,
                    Accent6 = A.ColorSchemeIndexValues.Accent6,
                    Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                },
                new TextStyles(
                    new TitleStyle(),
                    new BodyStyle(),
                    new OtherStyle()
                )
            );
        }

        private SlideLayout CreateSlideLayout()
        {
            return new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = 0, Cy = 0 },
                        new A.ChildOffset { X = 0, Y = 0 },
                        new A.ChildExtents { Cx = 12192000, Cy = 6858000 }
                    ))
                )),
                new P.ColorMapOverride(new A.MasterColorMapping())
            ) { Type = SlideLayoutValues.Blank };
        }
    }
}
