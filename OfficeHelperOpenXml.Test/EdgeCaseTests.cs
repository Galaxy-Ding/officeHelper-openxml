using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Core.Writers;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Test
{
    public class EdgeCaseTests
    {
        private SlideWriter CreateSlideWriter()
        {
            return new SlideWriter(new RelationshipIdManager());
        }

        [Fact]
        public void TestShapeWithNoText_CreatesEmptyParagraph()
        {
            // Arrange
            var slideWriter = CreateSlideWriter();
            var shapeData = new ShapeJsonData
            {
                Type = "textbox",
                Name = "NoTextShape",
                Box = "1.0,1.0,5.0,3.0",
                HasText = 0, // No text
                Text = null
            };

            // Act
            var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;

            // Assert
            Assert.NotNull(shape);
            Assert.NotNull(shape.TextBody);
            var paragraphs = shape.TextBody.Elements<A.Paragraph>();
            Assert.Single(paragraphs); // Should have one empty paragraph
        }

        [Fact]
        public void TestShapeWithZeroDimensions_AppliesMinimalSize()
        {
            // Arrange
            var slideWriter = CreateSlideWriter();
            var shapeData = new ShapeJsonData
            {
                Type = "textbox",
                Name = "ZeroSizeShape",
                Box = "0.00,0.00,0.00,0.00", // Zero dimensions
                HasText = 0
            };

            // Act
            var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;

            // Assert
            Assert.NotNull(shape);
            Assert.NotNull(shape.ShapeProperties);
            var transform = shape.ShapeProperties.Transform2D;
            Assert.NotNull(transform);
            Assert.NotNull(transform.Extents);
            // Should have minimal dimensions, not zero
            Assert.True(transform.Extents.Cx > 0);
            Assert.True(transform.Extents.Cy > 0);
        }

        [Fact]
        public void TestTransparentFill_CreatesNoFill()
        {
            // Arrange
            var slideWriter = CreateSlideWriter();
            var shapeData = new ShapeJsonData
            {
                Type = "textbox",
                Name = "TransparentShape",
                Box = "1.0,1.0,5.0,3.0",
                Fill = new FillJsonData
                {
                    Color = "RGB(255,0,0)",
                    Opacity = 0.0f // Transparent
                },
                HasText = 0
            };

            // Act
            var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;

            // Assert
            Assert.NotNull(shape);
            Assert.NotNull(shape.ShapeProperties);
            var noFill = shape.ShapeProperties.GetFirstChild<A.NoFill>();
            Assert.NotNull(noFill); // Should have NoFill element
        }

        [Fact]
        public void TestSolidFill_CreatesSolidFill()
        {
            // Arrange
            var slideWriter = CreateSlideWriter();
            var shapeData = new ShapeJsonData
            {
                Type = "textbox",
                Name = "SolidShape",
                Box = "1.0,1.0,5.0,3.0",
                Fill = new FillJsonData
                {
                    Color = "RGB(255,0,0)",
                    Opacity = 1.0f // Solid
                },
                HasText = 0
            };

            // Act
            var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;

            // Assert
            Assert.NotNull(shape);
            Assert.NotNull(shape.ShapeProperties);
            var solidFill = shape.ShapeProperties.GetFirstChild<A.SolidFill>();
            Assert.NotNull(solidFill); // Should have SolidFill element
        }

        [Fact]
        public void TestNoOutline_CreatesOutlineWithNoFill()
        {
            // Arrange
            var slideWriter = CreateSlideWriter();
            var shapeData = new ShapeJsonData
            {
                Type = "textbox",
                Name = "NoOutlineShape",
                Box = "1.0,1.0,5.0,3.0",
                Line = new LineJsonData
                {
                    HasOutline = 0 // No outline
                },
                HasText = 0
            };

            // Act
            var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;

            // Assert
            Assert.NotNull(shape);
            Assert.NotNull(shape.ShapeProperties);
            var outline = shape.ShapeProperties.GetFirstChild<A.Outline>();
            Assert.NotNull(outline);
            var noFill = outline.GetFirstChild<A.NoFill>();
            Assert.NotNull(noFill); // Outline should have NoFill
        }

        [Fact]
        public void TestNoShadow_DoesNotAddShadow()
        {
            // Arrange
            var slideWriter = CreateSlideWriter();
            var shapeData = new ShapeJsonData
            {
                Type = "textbox",
                Name = "NoShadowShape",
                Box = "1.0,1.0,5.0,3.0",
                Shadow = new ShadowJsonData
                {
                    HasShadow = 0 // No shadow
                },
                HasText = 0
            };

            // Act
            var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;

            // Assert
            Assert.NotNull(shape);
            Assert.NotNull(shape.ShapeProperties);
            var effectStyle = shape.ShapeProperties.GetFirstChild<A.EffectStyle>();
            Assert.Null(effectStyle); // Should not have effect style when no shadow
        }

        [Fact]
        public void TestEdgeCaseIntegration_ConvertsSuccessfully()
        {
            // Arrange
            var testData = new PresentationJsonData
            {
                MasterSlides = new System.Collections.Generic.List<SlideJsonData>(),
                ContentSlides = new System.Collections.Generic.List<SlideJsonData>
                {
                    new SlideJsonData
                    {
                        PageNumber = 1,
                        Title = "Edge Case Test",
                        Shapes = new System.Collections.Generic.List<ShapeJsonData>
                        {
                            // Shape with no text
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "NoText",
                                Box = "1.0,1.0,5.0,3.0",
                                HasText = 0
                            },
                            // Shape with zero dimensions
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "ZeroSize",
                                Box = "0.00,0.00,0.00,0.00",
                                HasText = 0
                            },
                            // Shape with transparent fill
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "Transparent",
                                Box = "6.0,1.0,5.0,3.0",
                                Fill = new FillJsonData { Opacity = 0.0f },
                                HasText = 0
                            },
                            // Shape with solid fill
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "Solid",
                                Box = "11.0,1.0,5.0,3.0",
                                Fill = new FillJsonData { Color = "RGB(0,255,0)", Opacity = 1.0f },
                                HasText = 0
                            },
                            // Shape with no outline
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "NoOutline",
                                Box = "1.0,5.0,5.0,3.0",
                                Line = new LineJsonData { HasOutline = 0 },
                                HasText = 0
                            },
                            // Shape with no shadow
                            new ShapeJsonData
                            {
                                Type = "textbox",
                                Name = "NoShadow",
                                Box = "6.0,5.0,5.0,3.0",
                                Shadow = new ShadowJsonData { HasShadow = 0 },
                                HasText = 0
                            }
                        }
                    }
                }
            };

            var jsonPath = "test_ppt/edge_case_test.json";
            var outputPath = "test_ppt/edge_case_test.pptx";

            // Create JSON file
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(testData, Formatting.Indented));

            var converter = new JsonToPptxConverter();

            // Act
            var result = converter.Convert(jsonPath, outputPath);

            // Assert
            Assert.True(result, "Conversion should succeed");
            Assert.True(File.Exists(outputPath), "Output file should exist");

            // Verify the PPTX can be opened
            using (var doc = PresentationDocument.Open(outputPath, false))
            {
                Assert.NotNull(doc.PresentationPart);
                Assert.NotNull(doc.PresentationPart.Presentation);
                var slideIdList = doc.PresentationPart.Presentation.SlideIdList;
                Assert.NotNull(slideIdList);
                Assert.Single(slideIdList.Elements<P.SlideId>()); // Should have 1 slide
            }

            // Cleanup
            if (File.Exists(jsonPath)) File.Delete(jsonPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }
}
