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
    /// Property-based tests for element removal (spLocks, ph elements)
    /// Feature: pptx-phase2-implementation-fixes, Property 9
    /// </summary>
    public class ElementRemovalPropertyTests
    {
        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 9: Non-placeholder shapes omit ph elements
        /// Validates: Requirements 10.1, 10.2, 10.3, 10.4, 10.5
        /// 
        /// Property: For any generated non-placeholder shape, the shape should not contain PlaceholderShape elements.
        /// Additionally, non-locked shapes should not contain ShapeLocks elements.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property NonPlaceholderShapesOmitPhElements()
        {
            return Prop.ForAll(
                GenerateNonPlaceholderShapeData(),
                shapeData =>
                {
                    // Skip null shape data
                    if (shapeData == null)
                        return true.ToProperty().Label("Null shape data skipped");
                    
                    try
                    {
                        // Create a SlideWriter instance
                        var relationshipIdManager = new RelationshipIdManager();
                        var slideWriter = new SlideWriter(relationshipIdManager);
                        
                        // Create the shape from JSON
                        var shape = slideWriter.CreateShapeFromJson(shapeData, 1, null) as P.Shape;
                        
                        if (shape == null)
                            return false.ToProperty().Label("FAIL: CreateShapeFromJson returned null");
                        
                        // Verify PlaceholderShape element is NOT present
                        var nvSpProps = shape.NonVisualShapeProperties;
                        if (nvSpProps == null)
                            return false.ToProperty().Label("FAIL: NonVisualShapeProperties is null");
                        
                        var appNvDrawingProps = nvSpProps.ApplicationNonVisualDrawingProperties;
                        if (appNvDrawingProps == null)
                            return false.ToProperty().Label("FAIL: ApplicationNonVisualDrawingProperties is null");
                        
                        var placeholderShape = appNvDrawingProps.GetFirstChild<PlaceholderShape>();
                        if (placeholderShape != null)
                            return false.ToProperty().Label("FAIL: PlaceholderShape element found in non-placeholder shape");
                        
                        // Verify ShapeLocks element is NOT present (for non-locked shapes)
                        var nvSpDrawingProps = nvSpProps.NonVisualShapeDrawingProperties;
                        if (nvSpDrawingProps == null)
                            return false.ToProperty().Label("FAIL: NonVisualShapeDrawingProperties is null");
                        
                        var shapeLocks = nvSpDrawingProps.GetFirstChild<A.ShapeLocks>();
                        if (shapeLocks != null)
                            return false.ToProperty().Label("FAIL: ShapeLocks element found in non-locked shape");
                        
                        return true.ToProperty().Label($"PASS: No unnecessary elements in {shapeData.Type} shape");
                    }
                    catch (Exception ex)
                    {
                        return false.ToProperty().Label($"FAIL: Exception during shape creation: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Feature: pptx-phase2-implementation-fixes, Property 9: Element removal in content slides
        /// Validates: Requirements 10.3
        /// 
        /// Property: For any content slide with non-placeholder shapes, all shapes should omit ph and spLocks elements.
        /// </summary>
        [Property(MaxTest = 50)]
        public Property ContentSlideShapesOmitUnnecessaryElements()
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
                                
                                // Check each shape for unnecessary elements
                                foreach (var shape in shapes)
                                {
                                    var nvSpProps = shape.NonVisualShapeProperties;
                                    if (nvSpProps == null)
                                        return false.ToProperty().Label("FAIL: NonVisualShapeProperties is null");
                                    
                                    // Check for PlaceholderShape
                                    var appNvDrawingProps = nvSpProps.ApplicationNonVisualDrawingProperties;
                                    if (appNvDrawingProps != null)
                                    {
                                        var placeholderShape = appNvDrawingProps.GetFirstChild<PlaceholderShape>();
                                        if (placeholderShape != null)
                                            return false.ToProperty().Label("FAIL: Content slide shape has PlaceholderShape element");
                                    }
                                    
                                    // Check for ShapeLocks
                                    var nvSpDrawingProps = nvSpProps.NonVisualShapeDrawingProperties;
                                    if (nvSpDrawingProps != null)
                                    {
                                        var shapeLocks = nvSpDrawingProps.GetFirstChild<A.ShapeLocks>();
                                        if (shapeLocks != null)
                                            return false.ToProperty().Label("FAIL: Content slide shape has ShapeLocks element");
                                    }
                                }
                                
                                return true.ToProperty().Label($"PASS: All {shapes.Count()} content slide shapes omit unnecessary elements");
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
        /// Generator for non-placeholder ShapeJsonData with random properties
        /// </summary>
        private static Arbitrary<ShapeJsonData> GenerateNonPlaceholderShapeData()
        {
            var shapeGen = from shapeType in Gen.Elements("textbox", "autoshape", "rectangle")
                          from name in Gen.Elements("Shape1", "TextBox1", "Content", "CustomShape")
                          from left in Gen.Choose(0, 20).Select(x => (float)x)
                          from top in Gen.Choose(0, 15).Select(x => (float)x)
                          from width in Gen.Choose(5, 20).Select(x => (float)x)
                          from height in Gen.Choose(2, 10).Select(x => (float)x)
                          from hasText in Gen.Elements(0, 1)
                          from textContent in Gen.Elements("Hello", "World", "Test", "Sample", "Content")
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
        /// Generator for SlideJsonData with random non-placeholder shapes
        /// </summary>
        private static Arbitrary<SlideJsonData> GenerateContentSlideData()
        {
            var slideGen = from pageNum in Gen.Choose(1, 10)
                          from title in Gen.Elements("Slide 1", "Slide 2", "Test Slide", "Content Slide")
                          from shapeCount in Gen.Choose(1, 3)
                          from shapes in Gen.ListOf(shapeCount, GenerateNonPlaceholderShapeData().Generator)
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
