using System;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Elements
{
    public class PictureElement : BaseElement
    {
        public override string ElementType => "Picture";
        
        public ShapeBounds Bounds { get; set; }
        
        public PictureElement() : base()
        {
            Bounds = new ShapeBounds();
        }
        
        protected override void InitializeComponents(Shape shape, SlidePart slidePart)
        {
        }
        
        public void InitializeFromPicture(Picture picture, SlidePart slidePart)
        {
            InitializeFromPictureInternal(picture, slidePart, null, null);
        }
        
        // 重载：从母版中初始化图片
        public void InitializeFromPicture(Picture picture, SlideMasterPart masterPart)
        {
            InitializeFromPictureInternal(picture, null, masterPart, null);
        }
        
        // 重载：从布局中初始化图片
        public void InitializeFromPicture(Picture picture, SlideLayoutPart layoutPart)
        {
            InitializeFromPictureInternal(picture, null, null, layoutPart);
        }
        
        // 内部通用方法
        private void InitializeFromPictureInternal(Picture picture, 
                                                   SlidePart slidePart = null, 
                                                   SlideMasterPart masterPart = null, 
                                                   SlideLayoutPart layoutPart = null)
        {
            var nvPicPr = picture.NonVisualPictureProperties;
            if (nvPicPr?.NonVisualDrawingProperties != null)
            {
                Name = nvPicPr.NonVisualDrawingProperties.Name ?? "";
            }
            
            ExtractBounds(picture);
            
            // 根据不同的 Part 类型提取图片内容
            var pictureComponent = new PictureComponent();
            if (slidePart != null)
            {
                pictureComponent.ExtractFromPicture(picture, slidePart);
            }
            else if (masterPart != null)
            {
                pictureComponent.ExtractFromPicture(picture, masterPart);
            }
            else if (layoutPart != null)
            {
                pictureComponent.ExtractFromPicture(picture, layoutPart);
            }
            
            if (pictureComponent.HasPicture)
            {
                AddComponent(pictureComponent);
            }
            
            var lineComponent = new LineComponent();
            var spPr = picture.ShapeProperties;
            if (spPr != null)
            {
                var tempShape = new Shape();
                tempShape.ShapeProperties = (ShapeProperties)spPr.CloneNode(true);
                lineComponent.ExtractFromShape(tempShape, slidePart);
            }
            AddComponent(lineComponent);
            
            // Extract shadow
            var shadowComponent = new ShadowComponent();
            if (spPr != null)
            {
                var tempShape = new Shape();
                tempShape.ShapeProperties = (ShapeProperties)spPr.CloneNode(true);
                shadowComponent.ExtractFromShape(tempShape, slidePart);
            }
            AddComponent(shadowComponent);
        }
        
        private void ExtractBounds(Picture picture)
        {
            try
            {
                var spPr = picture.ShapeProperties;
                if (spPr == null) return;
                
                var xfrm = spPr.Transform2D;
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
                Console.WriteLine($"提取图片位置信息时出错: {ex.Message}");
            }
        }
        
        public override string ToJson()
        {
            var parts = new System.Collections.Generic.List<string>();
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
            
            foreach (var c in GetAllComponents())
            {
                if (!c.IsEnabled) continue;
                try
                {
                    var json = c.ToJson();
                    if (json != "null")
                    {
                        var content = json.Trim();
                        if (c.ComponentType == "Picture")
                        {
                            if (content.StartsWith("{") && content.EndsWith("}"))
                                content = content.Substring(1, content.Length - 2);
                            parts.Add($"\"picture\":{{{content}}}");
                        }
                        else if (c.ComponentType == "Fill")
                        {
                            if (content.StartsWith("{") && content.EndsWith("}"))
                                content = content.Substring(1, content.Length - 2);
                            parts.Add($"\"fill\":{{{content}}}");
                        }
                        else if (c.ComponentType == "Line")
                        {
                            if (content.StartsWith("{") && content.EndsWith("}"))
                                content = content.Substring(1, content.Length - 2);
                            parts.Add($"\"line\":{{{content}}}");
                        }
                        else if (c.ComponentType == "Shadow")
                        {
                            if (content.StartsWith("{") && content.EndsWith("}"))
                                content = content.Substring(1, content.Length - 2);
                            parts.Add($"\"shadow\":{{{content}}}");
                        }
                    }
                }
                catch { }
            }
            
            return "{" + string.Join(",", parts) + "}";
        }
    }
}
