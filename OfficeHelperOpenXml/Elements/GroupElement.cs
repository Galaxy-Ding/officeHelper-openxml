using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Elements
{
    public class GroupElement : BaseElement
    {
        public override string ElementType => "Group";
        
        public ShapeBounds Bounds { get; set; }
        public List<BaseElement> Children { get; set; }
        
        public GroupElement() : base()
        {
            Bounds = new ShapeBounds();
            Children = new List<BaseElement>();
        }
        
        protected override void InitializeComponents(Shape shape, SlidePart slidePart) { }
        
        public void InitializeFromGroupShape(GroupShape groupShape, SlidePart slidePart)
        {
            var nvGrpSpPr = groupShape.NonVisualGroupShapeProperties;
            if (nvGrpSpPr?.NonVisualDrawingProperties != null)
            {
                Name = nvGrpSpPr.NonVisualDrawingProperties.Name ?? "";
            }
            
            ExtractBounds(groupShape);
            ExtractChildren(groupShape, slidePart);
        }
        
        private void ExtractBounds(GroupShape groupShape)
        {
            try
            {
                var grpSpPr = groupShape.GroupShapeProperties;
                if (grpSpPr == null) return;
                
                var xfrm = grpSpPr.TransformGroup;
                if (xfrm == null) return;
                
                if (xfrm.Offset != null)
                {
                    Bounds.Left = (float)UnitConverter.EmuToCm(xfrm.Offset.X?.Value ?? 0);
                    Bounds.Top = (float)UnitConverter.EmuToCm(xfrm.Offset.Y?.Value ?? 0);
                }
                
                if (xfrm.Extents != null)
                {
                    Bounds.Width = (float)UnitConverter.EmuToCm(xfrm.Extents.Cx?.Value ?? 0);
                    Bounds.Height = (float)UnitConverter.EmuToCm(xfrm.Extents.Cy?.Value ?? 0);
                }
                
                if (xfrm.Rotation != null)
                {
                    Bounds.Rotation = xfrm.Rotation.Value / 60000f;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取组位置信息时出错: {ex.Message}");
            }
        }
        
        private void ExtractChildren(GroupShape groupShape, SlidePart slidePart)
        {
            Children = new List<BaseElement>();
            
            foreach (var child in groupShape.ChildElements)
            {
                try
                {
                    if (child is Shape shape)
                    {
                        var element = ElementFactory.CreateFromShape(shape, slidePart);
                        if (element != null) Children.Add(element);
                    }
                    else if (child is Picture picture)
                    {
                        var pictureElement = new PictureElement();
                        pictureElement.InitializeFromPicture(picture, slidePart);
                        Children.Add(pictureElement);
                    }
                    else if (child is GroupShape nestedGroup)
                    {
                        var groupElement = new GroupElement();
                        groupElement.InitializeFromGroupShape(nestedGroup, slidePart);
                        Children.Add(groupElement);
                    }
                    else if (child is ConnectionShape connShape)
                    {
                        var connElement = new ConnectionElement();
                        connElement.InitializeFromConnectionShape(connShape, slidePart);
                        Children.Add(connElement);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"提取组子元素时出错: {ex.Message}");
                }
            }
        }
        
        public override string ToJson()
        {
            var parts = new List<string>();
            parts.Add($"\"type\":\"{ElementType.ToLower()}\"");
            parts.Add($"\"name\":\"{Name}\"");
            
            if (!string.IsNullOrEmpty(SpecialType))
            {
                parts.Add($"\"special_type\":\"{SpecialType}\"");
            }
            
            string boxStr = $"{Bounds.Left:F2},{Bounds.Top:F2},{Bounds.Width:F2},{Bounds.Height:F2}";
            parts.Add($"\"box\":\"{boxStr}\"");
            parts.Add($"\"rotation\":{Bounds.Rotation:F1}");
            parts.Add($"\"hastext\":0");
            
            var childrenJson = new List<string>();
            foreach (var child in Children)
            {
                try { childrenJson.Add(child.ToJson()); }
                catch { }
            }
            parts.Add($"\"children\":[{string.Join(",", childrenJson)}]");
            
            return "{" + string.Join(",", parts) + "}";
        }
    }
}
