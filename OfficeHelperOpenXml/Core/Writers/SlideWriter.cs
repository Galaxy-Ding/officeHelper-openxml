using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Elements;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;
using OfficeHelperOpenXml.Core.Converters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Core.Writers
{
    public class SlideWriter
    {
        private const int EMU_PER_CM = 360000;
        private readonly RelationshipIdManager _relationshipIdManager;

        /// <summary>
        /// Initializes a new instance of SlideWriter with a RelationshipIdManager
        /// </summary>
        /// <param name="relationshipIdManager">The relationship ID manager for generating sequential IDs</param>
        /// <exception cref="ArgumentNullException">Thrown when relationshipIdManager is null</exception>
        public SlideWriter(RelationshipIdManager relationshipIdManager)
        {
            _relationshipIdManager = relationshipIdManager ?? throw new ArgumentNullException(nameof(relationshipIdManager));
        }

        public bool AddElement(SlidePart slidePart, IElement element)
        {
            if (slidePart?.Slide?.CommonSlideData?.ShapeTree == null || element == null)
                return false;
            try
            {
                var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
                uint nextId = GetNextShapeId(shapeTree);
                OpenXmlElement shape = CreateShape(element, nextId, slidePart);
                if (shape != null) { shapeTree.Append(shape); return true; }
                return false;
            }
            catch { return false; }
        }

        public bool UpdateElement(SlidePart slidePart, int elementIndex, IElement element)
        {
            if (slidePart?.Slide?.CommonSlideData?.ShapeTree == null || element == null)
                return false;
            try
            {
                var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
                var shapes = shapeTree.Elements<P.Shape>().ToList();
                if (elementIndex < 0 || elementIndex >= shapes.Count) return false;
                var oldShape = shapes[elementIndex];
                uint id = oldShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id ?? GetNextShapeId(shapeTree);
                var newShape = CreateShape(element, id, slidePart) as P.Shape;
                if (newShape != null) { shapeTree.InsertAfter(newShape, oldShape); oldShape.Remove(); return true; }
                return false;
            }
            catch { return false; }
        }

        public bool DeleteElement(SlidePart slidePart, int elementIndex)
        {
            if (slidePart?.Slide?.CommonSlideData?.ShapeTree == null) return false;
            try
            {
                var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
                var shapes = shapeTree.Elements<P.Shape>().ToList();
                if (elementIndex < 0 || elementIndex >= shapes.Count) return false;
                shapes[elementIndex].Remove();
                return true;
            }
            catch { return false; }
        }

        private uint GetNextShapeId(ShapeTree shapeTree)
        {
            uint maxId = 1;
            foreach (var child in shapeTree.ChildElements)
            {
                uint? id = null;
                if (child is P.Shape shape) id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                else if (child is P.Picture pic) id = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (id.HasValue && id.Value > maxId) maxId = id.Value;
            }
            return maxId + 1;
        }

        private OpenXmlElement CreateShape(IElement element, uint id, SlidePart slidePart)
        {
            var posComp = element.GetComponent<PositionComponent>();
            var textComp = element.GetComponent<TextComponent>();
            var fillComp = element.GetComponent<FillComponent>();
            var lineComp = element.GetComponent<LineComponent>();
            float left = posComp?.Bounds?.X ?? 0;
            float top = posComp?.Bounds?.Y ?? 0;
            float width = posComp?.Bounds?.Width ?? 100;
            float height = posComp?.Bounds?.Height ?? 100;
            switch (element.ElementType.ToLower())
            {
                case "autoshape":
                case "textbox":
                    return CreateAutoShape(id, element.Name, left, top, width, height, textComp, fillComp, lineComp);
                case "picture":
                    var picComp = element.GetComponent<PictureComponent>();
                    return CreatePicture(id, element.Name, left, top, width, height, picComp, slidePart);
                default:
                    return CreateAutoShape(id, element.Name, left, top, width, height, textComp, fillComp, lineComp);
            }
        }

        private P.Shape CreateAutoShape(uint id, string name, float left, float top, float width, float height,
            TextComponent textComp, FillComponent fillComp, LineComponent lineComp)
        {
            var shape = new P.Shape();
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name ?? $"Shape {id}" },
                new P.NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            );
            shape.ShapeProperties = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = (long)(left * EMU_PER_CM), Y = (long)(top * EMU_PER_CM) },
                    new A.Extents { Cx = (long)(width * EMU_PER_CM), Cy = (long)(height * EMU_PER_CM) }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );
            if (fillComp != null && fillComp.Fill.HasFill)
            {
                var color = ParseColor(fillComp.Fill.Color?.ToString());
                shape.ShapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            }
            else { shape.ShapeProperties.Append(new A.NoFill()); }
            if (lineComp != null && lineComp.Line.HasOutline)
            {
                var color = ParseColor(lineComp.Line.Color?.ToString());
                shape.ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = color })) { Width = (int)(lineComp.Line.Weight * 12700) });
            }
            else { shape.ShapeProperties.Append(new A.Outline(new A.NoFill())); }
            shape.TextBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            if (textComp != null && textComp.Paragraphs.Count > 0)
            {
                foreach (var para in textComp.Paragraphs)
                {
                    var paragraph = new A.Paragraph();
                    // Add paragraph properties to disable bullets
                    var paragraphProperties = new A.ParagraphProperties();
                    paragraphProperties.Append(new A.NoBullet());
                    paragraph.Append(paragraphProperties);
                    
                    foreach (var run in para.Runs)
                    {
                        var aRun = new A.Run(
                            new A.RunProperties { FontSize = (int)(run.FontSize * 100), Bold = run.IsBold, Italic = run.IsItalic },
                            new A.Text(run.Text)
                        );
                        paragraph.Append(aRun);
                    }
                    // Add end paragraph run properties
                    paragraph.Append(new A.EndParagraphRunProperties());
                    shape.TextBody.Append(paragraph);
                }
            }
            else { 
                // Empty paragraph for shapes without text - also disable bullets
                var emptyParagraph = new A.Paragraph();
                var paragraphProperties = new A.ParagraphProperties();
                paragraphProperties.Append(new A.NoBullet());
                emptyParagraph.Append(paragraphProperties);
                emptyParagraph.Append(new A.EndParagraphRunProperties());
                shape.TextBody.Append(emptyParagraph); 
            }
            return shape;
        }

        private P.Picture CreatePicture(uint id, string name, float left, float top, float width, float height,
            PictureComponent picComp, SlidePart slidePart)
        {
            if (picComp == null || string.IsNullOrEmpty(picComp.ImageBase64)) return null;
            try
            {
                byte[] imageBytes = Convert.FromBase64String(picComp.ImageBase64);
                var contentType = GetImageContentType(picComp.ImageFormat);
                var imagePart = slidePart.AddNewPart<ImagePart>(contentType, _relationshipIdManager.GetNextId());
                using (var stream = new MemoryStream(imageBytes)) { imagePart.FeedData(stream); }
                string relId = slidePart.GetIdOfPart(imagePart);
                var picture = new P.Picture();
                picture.NonVisualPictureProperties = new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name ?? $"Picture {id}" },
                    new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                );
                picture.BlipFill = new P.BlipFill(new A.Blip { Embed = relId }, new A.Stretch(new A.FillRectangle()));
                picture.ShapeProperties = new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = (long)(left * EMU_PER_CM), Y = (long)(top * EMU_PER_CM) },
                        new A.Extents { Cx = (long)(width * EMU_PER_CM), Cy = (long)(height * EMU_PER_CM) }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                );
                return picture;
            }
            catch { return null; }
        }

        private string GetImageContentType(string format)
        {
            switch (format?.ToLower())
            {
                case "jpg":
                case "jpeg": return "image/jpeg";
                case "gif": return "image/gif";
                case "bmp": return "image/bmp";
                case "tiff": return "image/tiff";
                default: return "image/png";
            }
        }

        /// <summary>
        /// Parses a color string to hex format.
        /// Supports hex format (#RRGGBB or RRGGBB) and RGB format (RGB(r,g,b)).
        /// </summary>
        /// <param name="colorStr">Color string in hex or RGB format</param>
        /// <returns>Hex color string without # prefix, defaults to "000000" if invalid</returns>
        private string ParseColor(string colorStr)
        {
            if (string.IsNullOrEmpty(colorStr)) return "000000";
            if (colorStr.StartsWith("#")) return colorStr.Substring(1);
            if (colorStr.StartsWith("RGB(", StringComparison.OrdinalIgnoreCase))
            {
                // Use ColorHelper for consistent RGB parsing
                string hex = ColorHelper.ParseRgbToHex(colorStr);
                return string.IsNullOrEmpty(hex) ? "000000" : hex;
            }
            return colorStr.Replace("#", "");
        }

        /// <summary>
        /// Creates an RGB color with optional alpha transparency.
        /// Common helper to reduce code duplication in shadow creation.
        /// </summary>
        /// <param name="colorStr">Color string to parse</param>
        /// <param name="transparency">Transparency value (0.0 = opaque, 1.0 = fully transparent)</param>
        /// <returns>RgbColorModelHex with alpha applied if transparency > 0</returns>
        private A.RgbColorModelHex CreateRgbColorWithAlpha(string colorStr, float transparency)
        {
            var rgbColor = new A.RgbColorModelHex { Val = ParseColor(colorStr) };
            
            if (transparency > 0)
            {
                int alphaValue = (int)((1.0f - transparency) * 100000);
                rgbColor.Append(new A.Alpha { Val = alphaValue });
            }
            
            return rgbColor;
        }

        /// <summary>
        /// Calculates distance and angle from X and Y offsets.
        /// Common helper for shadow positioning.
        /// </summary>
        /// <param name="offsetX">X offset in points</param>
        /// <param name="offsetY">Y offset in points</param>
        /// <returns>Tuple of (distance in EMUs, angle in 60000ths of a degree)</returns>
        private (long distance, int angle) CalculateShadowDistanceAndAngle(float offsetX, float offsetY)
        {
            double distance = Math.Sqrt(offsetX * offsetX + offsetY * offsetY);
            long distanceEmu = (long)(distance * 12700);
            
            double angleRadians = Math.Atan2(offsetY, offsetX);
            double angleDegrees = angleRadians * (180 / Math.PI);
            int angle60k = (int)(angleDegrees * 60000);
            
            return (distanceEmu, angle60k);
        }

        // ===== JSON Conversion Methods =====

        /// <summary>
        /// Creates an OpenXML shape from JSON shape data.
        /// Supports textbox, autoshape, picture, table, and connection shape types.
        /// </summary>
        /// <param name="shapeData">The JSON shape data containing type, position, formatting, and content</param>
        /// <param name="id">Unique identifier for the shape</param>
        /// <param name="slidePart">The slide part to add the shape to (required for pictures, optional for other types)</param>
        /// <returns>An OpenXML element representing the shape, or null if creation fails</returns>
        /// <example>
        /// <code>
        /// var relationshipIdManager = new RelationshipIdManager();
        /// var slideWriter = new SlideWriter(relationshipIdManager);
        /// var shapeData = new ShapeJsonData 
        /// { 
        ///     Type = "textbox",
        ///     Name = "Title",
        ///     Box = "1.0,2.0,10.0,3.0",
        ///     HasText = 1,
        ///     Text = new List&lt;TextRunJsonData&gt; { ... }
        /// };
        /// var shape = slideWriter.CreateShapeFromJson(shapeData, 1, slidePart);
        /// </code>
        /// </example>
        public OpenXmlElement CreateShapeFromJson(ShapeJsonData shapeData, uint id, SlidePart slidePart)
        {
            if (shapeData == null) return null;

            // Parse box dimensions (format: "left,top,width,height" in cm)
            var boxParts = ParseBox(shapeData.Box);
            float left = boxParts.left;
            float top = boxParts.top;
            float width = boxParts.width;
            float height = boxParts.height;

            OpenXmlElement shape = null;

            switch (shapeData.Type?.ToLower())
            {
                case "textbox":
                    shape = CreateTextBoxFromJson(id, shapeData.Name, left, top, width, height, shapeData);
                    break;
                case "autoshape":
                    shape = CreateAutoShapeFromJson(id, shapeData.Name, left, top, width, height, shapeData);
                    break;
                case "picture":
                    shape = CreatePictureFromJson(id, shapeData.Name, left, top, width, height, shapeData, slidePart);
                    break;
                case "table":
                    shape = CreateTableFromJson(id, shapeData.Name, left, top, width, height, shapeData);
                    break;
                case "connection":
                    shape = CreateConnectionFromJson(id, shapeData.Name, left, top, width, height, shapeData);
                    break;
                default:
                    // Default to textbox
                    shape = CreateTextBoxFromJson(id, shapeData.Name, left, top, width, height, shapeData);
                    break;
            }

            return shape;
        }

        private (float left, float top, float width, float height) ParseBox(string box)
        {
            if (string.IsNullOrEmpty(box))
                return (0, 0, 0, 0);

            var parts = box.Split(',');
            if (parts.Length != 4)
                return (0, 0, 0, 0);

            float.TryParse(parts[0].Trim(), out float left);
            float.TryParse(parts[1].Trim(), out float top);
            float.TryParse(parts[2].Trim(), out float width);
            float.TryParse(parts[3].Trim(), out float height);

            // Handle zero dimensions - apply minimal default size to make shape visible but small
            // This prevents rendering issues with zero-sized shapes
            const float MIN_DIMENSION = 0.01f; // 0.01 cm minimum
            if (width <= 0) width = MIN_DIMENSION;
            if (height <= 0) height = MIN_DIMENSION;

            return (left, top, width, height);
        }

        /// <summary>
        /// Creates a Transform2D object with position, size, and optional rotation.
        /// Common helper to reduce code duplication across shape creation methods.
        /// </summary>
        /// <param name="left">Left position in centimeters</param>
        /// <param name="top">Top position in centimeters</param>
        /// <param name="width">Width in centimeters</param>
        /// <param name="height">Height in centimeters</param>
        /// <param name="rotation">Rotation in degrees (0 = no rotation)</param>
        /// <returns>Transform2D with position, size, and rotation applied</returns>
        private A.Transform2D CreateTransform(float left, float top, float width, float height, float rotation = 0)
        {
            var transform = new A.Transform2D(
                new A.Offset { X = (long)(left * EMU_PER_CM), Y = (long)(top * EMU_PER_CM) },
                new A.Extents { Cx = (long)(width * EMU_PER_CM), Cy = (long)(height * EMU_PER_CM) }
            );

            if (rotation != 0)
            {
                // Convert degrees to 60000ths of a degree
                transform.Rotation = (int)(rotation * 60000);
            }

            return transform;
        }

        private P.Shape CreateTextBoxFromJson(uint id, string name, float left, float top, float width, float height, ShapeJsonData shapeData)
        {
            var shape = new P.Shape();

            // Non-visual properties with required attributes
            // Only add ShapeLocks if needed (for now, we'll omit it for regular shapes)
            var nvSpDrawingProps = new P.NonVisualShapeDrawingProperties();
            nvSpDrawingProps.TextBox = true; // Add txBox="1" attribute for text boxes
            
            // Only add PlaceholderShape if this is actually a placeholder
            // For regular shapes, omit the PlaceholderShape element
            shape.NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name ?? $"TextBox {id}" },
                nvSpDrawingProps,
                new ApplicationNonVisualDrawingProperties()
            );

            // Shape properties with positioning and rotation using helper
            shape.ShapeProperties = new P.ShapeProperties(
                CreateTransform(left, top, width, height, shapeData.Rotation),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );

            // Apply fill
            this.ApplyFillFromJson(shape.ShapeProperties, shapeData.Fill);

            // Apply line - ApplyLineFromJson handles both HasOutline=0 (NoFill) and HasOutline=1 (with outline)
            if (shapeData.Line != null)
            {
                this.ApplyLineFromJson(shape.ShapeProperties, shapeData.Line);
            }

            // Apply shadow
            this.ApplyShadowFromJson(shape.ShapeProperties, shapeData.Shadow);

            // Text body with required attributes
            var bodyProps = new A.BodyProperties
            {
                Wrap = A.TextWrappingValues.None,  // Add wrap="none" attribute
                RightToLeftColumns = false          // Add rtlCol="0" attribute
            };
            
            shape.TextBody = new P.TextBody(
                bodyProps,
                new A.ListStyle()
            );

            // Handle shapes with no text (hastext=0) - create empty paragraph
            // Handle shapes with text (hastext=1) - populate with text runs
            if (shapeData.HasText == 1 && shapeData.Text != null && shapeData.Text.Count > 0)
            {
                var paragraph = this.CreateParagraphFromJson(shapeData.Text);
                shape.TextBody.Append(paragraph);
            }
            else
            {
                // Empty paragraph for shapes without text - also disable bullets
                var emptyParagraph = new A.Paragraph();
                var paragraphProperties = new A.ParagraphProperties();
                paragraphProperties.Append(new A.NoBullet());
                emptyParagraph.Append(paragraphProperties);
                emptyParagraph.Append(new A.EndParagraphRunProperties()); // Add endParaRPr
                shape.TextBody.Append(emptyParagraph);
            }

            return shape;
        }

        private P.Shape CreateAutoShapeFromJson(uint id, string name, float left, float top, float width, float height, ShapeJsonData shapeData)
        {
            // AutoShape is similar to TextBox but may have different geometry
            return CreateTextBoxFromJson(id, name, left, top, width, height, shapeData);
        }

        private P.Picture CreatePictureFromJson(uint id, string name, float left, float top, float width, float height, ShapeJsonData shapeData, SlidePart slidePart)
        {
            // For now, create a placeholder picture
            // Full implementation would require image data in JSON
            var picture = new P.Picture();

            picture.NonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name ?? $"Picture {id}" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties()
            );

            // Use helper for transform
            picture.ShapeProperties = new P.ShapeProperties(
                CreateTransform(left, top, width, height, shapeData?.Rotation ?? 0),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            );

            return picture;
        }

        private P.GraphicFrame CreateTableFromJson(uint id, string name, float left, float top, float width, float height, ShapeJsonData shapeData)
        {
            // Placeholder for table creation
            // Full implementation would require table data structure
            var graphicFrame = new P.GraphicFrame();

            graphicFrame.NonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name ?? $"Table {id}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            );

            // Use helper for transform (P.Transform uses same structure as A.Transform2D)
            var transform2D = CreateTransform(left, top, width, height, shapeData?.Rotation ?? 0);
            graphicFrame.Transform = new P.Transform(
                transform2D.Offset,
                transform2D.Extents
            );

            return graphicFrame;
        }

        private P.ConnectionShape CreateConnectionFromJson(uint id, string name, float left, float top, float width, float height, ShapeJsonData shapeData)
        {
            // Placeholder for connection shape creation
            var connShape = new P.ConnectionShape();

            connShape.NonVisualConnectionShapeProperties = new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name ?? $"Connection {id}" },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            );

            // Use helper for transform
            connShape.ShapeProperties = new P.ShapeProperties(
                CreateTransform(left, top, width, height, shapeData?.Rotation ?? 0),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Line }
            );

            return connShape;
        }

        /// <summary>
        /// Applies fill properties from JSON to shape properties.
        /// Handles transparent fills (opacity=0), solid fills (opacity=1), and theme colors with transforms.
        /// </summary>
        /// <param name="shapeProps">The shape properties to apply fill to</param>
        /// <param name="fillData">The JSON fill data containing color, opacity, and theme color information</param>
        /// <remarks>
        /// When fillData is null or opacity is 0, creates a shape with no fill.
        /// When schemeColor is specified, prioritizes theme color over RGB.
        /// Supports luminance modulation (lumMod) and offset (lumOff) for theme colors.
        /// </remarks>
        /// <example>
        /// <code>
        /// var relationshipIdManager = new RelationshipIdManager();
        /// var slideWriter = new SlideWriter(relationshipIdManager);
        /// var fillData = new FillJsonData 
        /// { 
        ///     Color = "RGB(255,0,0)",
        ///     Opacity = 0.5f,
        ///     SchemeColor = "accent1",
        ///     ColorTransforms = new ColorTransformJsonData { LumMod = 60000, LumOff = 40000 }
        /// };
        /// slideWriter.ApplyFillFromJson(shapeProperties, fillData);
        /// </code>
        /// </example>
        public void ApplyFillFromJson(P.ShapeProperties shapeProps, FillJsonData fillData)
        {
            if (shapeProps == null) return;

            // Handle transparent fills (opacity=0.0) - create shape with no fill
            if (fillData == null || fillData.Opacity == 0)
            {
                // No fill - shape is transparent
                shapeProps.Append(new A.NoFill());
                return;
            }

            // Extract color transform values
            int? lumMod = fillData.ColorTransforms?.LumMod;
            int? lumOff = fillData.ColorTransforms?.LumOff;

            // Create color from JSON (supports both RGB and theme colors)
            var color = ColorHelper.CreateColorFromJson(fillData.Color, fillData.SchemeColor, lumMod, lumOff);

            if (color == null)
            {
                shapeProps.Append(new A.NoFill());
                return;
            }

            var solidFill = new A.SolidFill();
            solidFill.Append((OpenXmlElement)color);

            // Handle solid fills (opacity=1.0) - no alpha modification needed
            // Apply opacity if not fully opaque (0 < opacity < 1.0)
            if (fillData.Opacity < 1.0f && fillData.Opacity > 0)
            {
                int alphaValue = (int)((1.0f - fillData.Opacity) * 100000);
                if (color is A.RgbColorModelHex rgbColor)
                {
                    rgbColor.Append(new A.Alpha { Val = 100000 - alphaValue });
                }
                else if (color is A.SchemeColor schemeColor)
                {
                    schemeColor.Append(new A.Alpha { Val = 100000 - alphaValue });
                }
            }
            // When opacity=1.0, the fill is solid with no transparency

            shapeProps.Append(solidFill);
        }

        /// <summary>
        /// Applies line (outline) properties from JSON to shape properties.
        /// Handles shapes with no outline (has_outline=0) and supports theme colors.
        /// </summary>
        /// <param name="shapeProps">The shape properties to apply line to</param>
        /// <param name="lineData">The JSON line data containing outline settings</param>
        /// <remarks>
        /// When lineData is null or has_outline=0, creates a shape without outline.
        /// Width is converted from points to EMUs (1 pt = 12700 EMUs).
        /// Supports both RGB colors and theme colors with scheme color references.
        /// </remarks>
        public void ApplyLineFromJson(P.ShapeProperties shapeProps, LineJsonData lineData)
        {
            if (shapeProps == null) return;

            // Handle shapes with no outline (has_outline=0) - create shape without outline
            if (lineData == null || lineData.HasOutline == 0)
            {
                // No outline - shape has no border
                shapeProps.Append(new A.Outline(new A.NoFill()));
                return;
            }

            var outline = new A.Outline();

            // Set width (convert from points to EMUs: 1 pt = 12700 EMUs)
            if (lineData.Width > 0)
            {
                outline.Width = (int)(lineData.Width * 12700);
            }

            // Create color from JSON (supports both RGB and theme colors)
            var color = ColorHelper.CreateColorFromJson(lineData.Color, lineData.SchemeColor, null, null);

            if (color != null)
            {
                var solidFill = new A.SolidFill();
                solidFill.Append((OpenXmlElement)color);
                outline.Append(solidFill);
            }
            else
            {
                outline.Append(new A.NoFill());
            }

            shapeProps.Append(outline);
        }

        /// <summary>
        /// Applies shadow properties from JSON to shape properties.
        /// Supports both inner and outer shadow types with full parameter control.
        /// </summary>
        /// <param name="shapeProps">The shape properties to apply shadow to</param>
        /// <param name="shadowData">The JSON shadow data containing shadow settings</param>
        /// <remarks>
        /// When shadowData is null or has_shadow=0, no shadow is applied.
        /// Supports inner and outer shadow types based on shadow_type property.
        /// All measurements (blur, distance) are converted from points to EMUs.
        /// Shadow angle is calculated from OffsetX and OffsetY values.
        /// </remarks>
        public void ApplyShadowFromJson(P.ShapeProperties shapeProps, ShadowJsonData shadowData)
        {
            if (shapeProps == null) return;

            // Handle shapes with no shadow (has_shadow=0) - create shape without shadow
            if (shadowData == null || shadowData.HasShadow == 0)
            {
                // No shadow - shape has no shadow effect
                return;
            }

            // Create effect list
            var effectList = new A.EffectList();

            // Determine shadow type
            bool isInnerShadow = shadowData.ShadowType?.ToLower() == "inner";

            if (isInnerShadow)
            {
                var innerShadow = new A.InnerShadow();

                // Set blur radius (convert from points to EMUs)
                if (shadowData.Blur > 0)
                {
                    innerShadow.BlurRadius = (long)(shadowData.Blur * 12700);
                }

                // Set distance and angle using helper
                if (shadowData.OffsetX != 0 || shadowData.OffsetY != 0)
                {
                    var (distance, angle) = CalculateShadowDistanceAndAngle(shadowData.OffsetX, shadowData.OffsetY);
                    innerShadow.Distance = distance;
                    innerShadow.Direction = angle;
                }

                // Set color with transparency using helper
                if (!string.IsNullOrEmpty(shadowData.Color))
                {
                    innerShadow.RgbColorModelHex = CreateRgbColorWithAlpha(shadowData.Color, shadowData.Transparency);
                }

                effectList.Append(innerShadow);
            }
            else
            {
                // Outer shadow
                var outerShadow = new A.OuterShadow();

                // Set blur radius (convert from points to EMUs)
                if (shadowData.Blur > 0)
                {
                    outerShadow.BlurRadius = (long)(shadowData.Blur * 12700);
                }

                // Set distance and angle using helper
                if (shadowData.OffsetX != 0 || shadowData.OffsetY != 0)
                {
                    var (distance, angle) = CalculateShadowDistanceAndAngle(shadowData.OffsetX, shadowData.OffsetY);
                    outerShadow.Distance = distance;
                    outerShadow.Direction = angle;
                }

                // Set color with transparency using helper
                if (!string.IsNullOrEmpty(shadowData.Color))
                {
                    outerShadow.RgbColorModelHex = CreateRgbColorWithAlpha(shadowData.Color, shadowData.Transparency);
                }

                effectList.Append(outerShadow);
            }

            shapeProps.Append(new A.EffectStyle(effectList));
        }

        /// <summary>
        /// Creates a paragraph with multiple text runs from JSON.
        /// Preserves the order of text runs as specified in the JSON array.
        /// </summary>
        /// <param name="textRuns">List of text run data from JSON</param>
        /// <returns>An OpenXML paragraph containing all text runs with their formatting</returns>
        /// <remarks>
        /// Each text run maintains its own formatting properties (font, size, color, effects).
        /// Empty or null text run lists result in an empty paragraph.
        /// </remarks>
        public A.Paragraph CreateParagraphFromJson(List<TextRunJsonData> textRuns)
        {
            var paragraph = new A.Paragraph();

            // Add paragraph properties to disable bullets
            // Without this, PowerPoint may apply default bullet formatting
            var paragraphProperties = new A.ParagraphProperties();
            paragraphProperties.Append(new A.NoBullet());
            paragraph.Append(paragraphProperties);

            if (textRuns != null && textRuns.Count > 0)
            {
                foreach (var runData in textRuns)
                {
                    var run = CreateRunFromJson(runData);
                    if (run != null)
                    {
                        paragraph.Append(run);
                    }
                }
            }

            // Add end paragraph run properties
            paragraph.Append(new A.EndParagraphRunProperties());

            return paragraph;
        }

        /// <summary>
        /// Creates a text run with formatting from JSON.
        /// Applies font properties, text formatting (bold, italic, underline, strikethrough),
        /// color (RGB or theme), and text effects.
        /// </summary>
        /// <param name="runData">The JSON text run data containing content and formatting</param>
        /// <returns>An OpenXML run with all formatting applied, or null if runData is null</returns>
        /// <remarks>
        /// Font size is converted from points to hundredths of a point.
        /// Supports both RGB colors and theme colors with luminance transforms.
        /// Text effects (shadow, glow, reflection, soft edge) are applied when has_effects=1.
        /// </remarks>
        public A.Run CreateRunFromJson(TextRunJsonData runData)
        {
            if (runData == null) return null;

            var runProps = new A.RunProperties();

            // Add language attributes
            runProps.Language = "zh-CN";
            runProps.AlternativeLanguage = "en-US";
            runProps.Dirty = false;
            // Note: SmtClean property doesn't exist in OpenXML SDK, skipping

            // Only add explicit formatting when it differs from theme defaults
            // Font size - only add if explicitly specified (> 0)
            if (runData.FontSize > 0)
            {
                runProps.FontSize = (int)(runData.FontSize * 100);
            }

            // Font family - only add if explicitly specified (non-empty)
            if (!string.IsNullOrEmpty(runData.Font))
            {
                runProps.Append(new A.LatinFont { Typeface = runData.Font });
            }

            // Bold - only add if explicitly set to true
            if (runData.FontBold == 1)
            {
                runProps.Bold = true;
            }

            // Italic - only add if explicitly set to true
            if (runData.FontItalic == 1)
            {
                runProps.Italic = true;
            }

            // Underline - only add if explicitly set to true
            if (runData.FontUnderline == 1)
            {
                runProps.Underline = A.TextUnderlineValues.Single;
            }

            // Strikethrough - only add if explicitly set to true
            if (runData.FontStrikethrough == 1)
            {
                runProps.Strike = A.TextStrikeValues.SingleStrike;
            }

            // Font color - only add if explicitly specified (non-empty)
            // This allows theme colors to be inherited when not specified
            if (!string.IsNullOrEmpty(runData.FontColor) || !string.IsNullOrEmpty(runData.SchemeColor))
            {
                int? lumMod = runData.ColorTransforms?.LumMod;
                int? lumOff = runData.ColorTransforms?.LumOff;
                var color = ColorHelper.CreateColorFromJson(runData.FontColor, runData.SchemeColor, lumMod, lumOff);
                
                if (color != null)
                {
                    var solidFill = new A.SolidFill();
                    solidFill.Append((OpenXmlElement)color);
                    runProps.Append(solidFill);
                }
            }

            // Apply text effects
            if (runData.TextEffects != null && runData.TextEffects.HasEffects == 1)
            {
                ApplyTextEffects(runProps, runData.TextEffects);
            }

            var run = new A.Run();
            run.RunProperties = runProps;
            run.Text = new A.Text(runData.Content ?? "");

            return run;
        }

        /// <summary>
        /// Applies text effects (shadow, glow, reflection, soft edge) to run properties.
        /// Each effect is applied independently based on its has_* flag.
        /// </summary>
        /// <param name="runProps">The run properties to apply effects to</param>
        /// <param name="effects">The JSON text effects data</param>
        /// <remarks>
        /// Supports four types of text effects:
        /// - Shadow: Outer shadow with blur, distance, angle, and transparency
        /// - Glow: Glow effect with default radius
        /// - Reflection: Reflection effect with standard parameters
        /// - Soft Edge: Soft edge effect with default radius
        /// Effects are only applied when their respective has_* flag is set to 1.
        /// </remarks>
        public void ApplyTextEffects(A.RunProperties runProps, TextEffectsJsonData effects)
        {
            if (runProps == null || effects == null) return;

            var effectList = new A.EffectList();
            bool hasAnyEffect = false;

            // Text shadow
            if (effects.Shadow != null && effects.Shadow.HasShadow == 1)
            {
                var outerShadow = new A.OuterShadow();

                // Set blur radius (convert from points to EMUs)
                if (effects.Shadow.Blur > 0)
                {
                    outerShadow.BlurRadius = (long)(effects.Shadow.Blur * 12700);
                }

                // Set distance (convert from points to EMUs)
                if (effects.Shadow.Distance > 0)
                {
                    outerShadow.Distance = (long)(effects.Shadow.Distance * 12700);
                }

                // Set angle (convert from degrees to 60000ths of a degree)
                if (effects.Shadow.Angle != 0)
                {
                    outerShadow.Direction = (int)(effects.Shadow.Angle * 60000);
                }

                // Set color with transparency using helper
                if (!string.IsNullOrEmpty(effects.Shadow.Color))
                {
                    outerShadow.RgbColorModelHex = CreateRgbColorWithAlpha(effects.Shadow.Color, effects.Shadow.Transparency);
                }

                effectList.Append(outerShadow);
                hasAnyEffect = true;
            }

            // Text glow
            if (effects.Glow != null && effects.Glow.HasGlow == 1)
            {
                var glow = new A.Glow { Radius = 50000 }; // Default glow radius
                glow.Append(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
                effectList.Append(glow);
                hasAnyEffect = true;
            }

            // Text reflection
            if (effects.Reflection != null && effects.Reflection.HasReflection == 1)
            {
                var reflection = new A.Reflection
                {
                    BlurRadius = 6350L,
                    StartOpacity = 100000,
                    StartPosition = 0,
                    EndAlpha = 0,
                    EndPosition = 100000,
                    Distance = 0,
                    Direction = 5400000,
                    FadeDirection = 5400000,
                    Alignment = A.RectangleAlignmentValues.Bottom,
                    RotateWithShape = true
                };
                effectList.Append(reflection);
                hasAnyEffect = true;
            }

            // Text soft edge
            if (effects.SoftEdge != null && effects.SoftEdge.HasSoftEdge == 1)
            {
                var softEdge = new A.SoftEdge { Radius = 50000 }; // Default soft edge radius
                effectList.Append(softEdge);
                hasAnyEffect = true;
            }

            if (hasAnyEffect)
            {
                runProps.Append(new A.EffectStyle(effectList));
            }
        }
    }
}
