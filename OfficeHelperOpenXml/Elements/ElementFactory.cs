using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Elements
{
    public static class ElementFactory
    {
        public static BaseElement CreateFromShape(Shape shape, SlidePart slidePart)
        {
            if (shape == null) return null;

            try
            {
                var shapeType = DetermineShapeType(shape);
                BaseElement element;

                if (shapeType == ShapeCategory.TextBox)
                    element = new TextBoxElement();
                else
                    element = new AutoShapeElement();

                element.InitializeFromShape(shape, slidePart);
                return element;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建元素时出错: {ex.Message}");
                return null;
            }
        }

        public static BaseElement CreateFromGraphicFrame(GraphicFrame graphicFrame, SlidePart slidePart)
        {
            if (graphicFrame == null) return null;

            try
            {
                var table = graphicFrame.Descendants<A.Table>().FirstOrDefault();
                if (table != null)
                {
                    var tableElement = new TableElement();
                    tableElement.ExtractFromGraphicFrame(graphicFrame, slidePart);
                    return tableElement;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建GraphicFrame元素时出错: {ex.Message}");
                return null;
            }
        }

        private static ShapeCategory DetermineShapeType(Shape shape)
        {
            var nvSpPr = shape.NonVisualShapeProperties;
            if (nvSpPr?.NonVisualDrawingProperties != null)
            {
                string name = nvSpPr.NonVisualDrawingProperties.Name?.Value ?? "";
                if (name.IndexOf("TextBox", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("文本框", StringComparison.OrdinalIgnoreCase) >= 0)
                    return ShapeCategory.TextBox;
            }

            var spPr = shape.ShapeProperties;
            if (spPr != null)
            {
                var prstGeom = spPr.GetFirstChild<A.PresetGeometry>();
                if (prstGeom != null && prstGeom.Preset != null)
                {
                    if (prstGeom.Preset.Value == A.ShapeTypeValues.Rectangle)
                    {
                        if (shape.TextBody != null && shape.TextBody.HasChildren)
                        {
                            var hasText = false;
                            foreach (var para in shape.TextBody.Elements<A.Paragraph>())
                            {
                                foreach (var run in para.Elements<A.Run>())
                                {
                                    var text = run.GetFirstChild<A.Text>();
                                    if (text != null && !string.IsNullOrEmpty(text.Text))
                                    {
                                        hasText = true;
                                        break;
                                    }
                                }
                                if (hasText) break;
                            }

                            if (hasText)
                            {
                                var noFill = spPr.GetFirstChild<A.NoFill>();
                                if (noFill != null)
                                    return ShapeCategory.TextBox;
                            }
                        }
                    }
                }
            }

            return ShapeCategory.AutoShape;
        }

        private enum ShapeCategory
        {
            AutoShape,
            TextBox,
            Picture,
            Group,
            Table
        }
    }
}
