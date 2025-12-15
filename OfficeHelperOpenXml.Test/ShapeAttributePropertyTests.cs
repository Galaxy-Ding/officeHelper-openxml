using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using FsCheck;
using FsCheck.Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Core.Writers;
using OfficeHelperOpenXml.Models.Json;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Property-based tests for shape attributes (txBox, wrap, rtlCol)
    /// Feature: pptx-phase2-implementation-fixes, Property 8
    /// </summary>
    public class ShapeAttributePropertyTests
    {
        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 8: All text boxes have required attributes
        /// Validates: Requirements 9.1, 9.2, 9.3, 9.4, 9.5
        /// 
        /// Property: For any text box shape created from JSON, the shape should have:
        /// - txBox="1" attribute in NonVisualShapeDrawingProperties
        /// - wrap="none" attribute in BodyProperties
        /// - rtlCol="0" attribute in BodyProperties
        /// </summary>
        [Property(MaxTest = 100)]
        public Property AllTextBoxesHaveRequiredAttributes()
        {
            return Prop.ForAll(
                GenerateShapeJsonData(),
                shapeData =>
                {
                    // Skip null shape data
                    if (shapeData == null)
                        return true.ToProperty().Label("Null shape data skipped");
                    
                    // Only test textbox and autoshape types (which use CreateTextBoxFromJson)
                    if (shapeData.Type?.ToLower() != "textbox" && shapeData.Type?.ToLower() != "autoshape")
                        return true.ToProperty().Label($"Skipped non-textbox type: {shapeData.Type}");
                    
                    try
                    {
                        // Create a SlideWriter instance
                        var relationshipIdManager = new RelationshipIdManager();
                        var slideWriter = new SlideWriter(relationshipIdManager);
                        
                        // Create the shape from JSON
                        var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;
                        
                        if (shape == null)
                            return false.ToProperty().Label("FAIL: CreateShapeFromJson returned null");
                        
                        // Verify txBox="1" attribute
                        var nvSpDrawingProps = shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties;
                        if (nvSpDrawingProps == null)
                            return false.ToProperty().Label("FAIL: NonVisualShapeDrawingProperties is null");
                        
                        if (nvSpDrawingProps.TextBox == null || !nvSpDrawingProps.TextBox.Value)
                            return false.ToProperty().Label("FAIL: txBox attribute is not set to true");
                        
                        // Verify wrap and rtlCol attributes
                        var textBody = shape.TextBody;
                        if (textBody == null)
                            return false.ToProperty().Label("FAIL: TextBody is null");
                        
                        var bodyProps = textBody.GetFirstChild<A.BodyProperties>();
                        if (bodyProps == null)
                            return false.ToProperty().Label("FAIL: BodyProperties is null");
                        
                        if (bodyProps.Wrap == null || bodyProps.Wrap.Value != A.TextWrappingValues.None)
                            return false.ToProperty().Label($"FAIL: wrap attribute is not 'none', got: {bodyProps.Wrap?.Value}");
                        
                        if (bodyProps.RightToLeftColumns == null || bodyProps.RightToLeftColumns.Value != false)
                            return false.ToProperty().Label($"FAIL: rtlCol attribute is not false, got: {bodyProps.RightToLeftColumns?.Value}");
                        
                        return true.ToProperty().Label($"PASS: All attributes present for {shapeData.Type}");
                    }
                    catch (Exception ex)
                    {
                        return false.ToProperty().Label($"FAIL: Exception during shape creation: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 8: Shape attributes in content slides
        /// Validates: Requirements 9.4
        /// 
        /// Property: For any content slide with text box shapes, all shapes should have required attributes.
        /// </summary>
        [Property(MaxTest = 50)]
        public Property ContentSlideShapesHaveRequiredAttributes()
        {
            return Prop.ForAll(
                GenerateContentSlideData(),
                slideData =>
                {
                    // Skip null or empty slide data
                    if (slideData == null || slideData.Shapes == null || slideData.Shapes.Count == 0)
                        return true.ToProperty().Label("No shapes to test");
                    
                    try
                    {
                        // Create temporary files for JSON and PPTX
                        var tempJsonFile = Path.GetTempFileName() + ".json";
                        var tempPptxFile = Path.GetTempFileName() + ".pptx";
                        
                        try
                        {
                            // Create presentation data with the content slide
                            var presentationData = new PresentationJsonData
                            {
                                ContentSlides = new List<SlideJsonData> { slideData }
                            };
                            
                            // Write JSON to temp file
                            var json = Newtonsoft.Json.JsonConvert.SerializeObject(presentationData, Newtonsoft.Json.Formatting.Indented);
                            File.WriteAllText(tempJsonFile, json);
                            
                            // Convert to PPTX
                            var converter = new JsonToPptxConverter();
                            var success = converter.Convert(tempJsonFile, tempPptxFile);
                            
                            if (!success)
                                return false.ToProperty().Label("FAIL: Conversion failed");
                            
                            // Open the generated PPTX and verify shapes
                            using (var document = PresentationDocument.Open(tempPptxFile, false))
                            {
                                var slidePart = document.PresentationPart?.SlideParts.FirstOrDefault();
                                if (slidePart == null)
                                    return false.ToProperty().Label("FAIL: No slide part found");
                                
                                var shapes = slidePart.Slide?.CommonSlideData?.ShapeTree?.Elements<P.Shape>();
                                if (shapes == null || !shapes.Any())
                                    return true.ToProperty().Label("No shapes in slide");
                                
                                // Check each shape for required attributes
                                foreach (var shape in shapes)
                                {
                                    var nvSpDrawingProps = shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties;
                                    if (nvSpDrawingProps?.TextBox == null || !nvSpDrawingProps.TextBox.Value)
                                        return false.ToProperty().Label("FAIL: Content slide shape missing txBox attribute");
                                    
                                    var bodyProps = shape.TextBody?.GetFirstChild<A.BodyProperties>();
                                    if (bodyProps == null)
                                        return false.ToProperty().Label("FAIL: Content slide shape missing BodyProperties");
                                    
                                    if (bodyProps.Wrap == null || bodyProps.Wrap.Value != A.TextWrappingValues.None)
                                        return false.ToProperty().Label("FAIL: Content slide shape missing wrap attribute");
                                    
                                    if (bodyProps.RightToLeftColumns == null || bodyProps.RightToLeftColumns.Value != false)
                                        return false.ToProperty().Label("FAIL: Content slide shape missing rtlCol attribute");
                                }
                                
                                return true.ToProperty().Label($"PASS: All {shapes.Count()} content slide shapes have required attributes");
                            }
                        }
                        finally
                        {
                            // Clean up temp files
                            if (File.Exists(tempJsonFile))
                                File.Delete(tempJsonFile);
                            if (File.Exists(tempPptxFile))
                                File.Delete(tempPptxFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        return false.ToProperty().Label($"FAIL: Exception: {ex.Message}");
                    }
                });
        }
        
        /// <summary>
        /// Generator for ShapeJsonData with random properties
        /// </summary>
        private static Arbitrary<ShapeJsonData> GenerateShapeJsonData()
        {
            var shapeGen = from shapeType in Gen.Elements("textbox", "autoshape")
                          from name in Gen.Elements("Shape1", "TextBox1", "Title", "Content")
                          from left in Gen.Choose(0, 20).Select(x => (float)x)
                          from top in Gen.Choose(0, 15).Select(x => (float)x)
                          from width in Gen.Choose(5, 20).Select(x => (float)x)
                          from height in Gen.Choose(2, 10).Select(x => (float)x)
                          from hasText in Gen.Elements(0, 1)
                          from textContent in Gen.Elements("Hello", "World", "Test", "Sample")
                          select new ShapeJsonData
                          {
                              Type = shapeType,
                              Name = name,
                              Box = $"{left},{top},{width},{height}",
                              HasText = hasText,
                              Text = hasText == 1 ? new List<TextRunJsonData>
                              {
                                  new TextRunJsonData
                                  {
                                      Content = textContent,
                                      Font = "Arial",
                                      FontSize = 12,
                                      FontColor = "RGB(0,0,0)",
                                      FontBold = 0,
                                      FontItalic = 0,
                                      FontUnderline = 0,
                                      FontStrikethrough = 0
                                  }
                              } : new List<TextRunJsonData>(),
                              Fill = new FillJsonData
                              {
                                  Color = "RGB(255,255,255)",
                                  Opacity = 1.0f
                              },
                              Line = new LineJsonData
                              {
                                  HasOutline = 1,
                                  Color = "RGB(0,0,0)",
                                  Width = 1.0f
                              },
                              Rotation = 0
                          };
            
            return Arb.From(shapeGen);
        }
        
        /// <summary>
        /// Generator for SlideJsonData with random shapes
        /// </summary>
        private static Arbitrary<SlideJsonData> GenerateContentSlideData()
        {
            var slideGen = from pageNum in Gen.Choose(1, 10)
                          from title in Gen.Elements("Slide 1", "Slide 2", "Test Slide")
                          from shapeCount in Gen.Choose(1, 3)
                          from shapes in Gen.ListOf(shapeCount, GenerateShapeJsonData().Generator)
                          select new SlideJsonData
                          {
                              PageNumber = pageNum,
                              Title = title,
                              Shapes = shapes.ToList()
                          };
            
            return Arb.From(slideGen);
        }
    }
}
