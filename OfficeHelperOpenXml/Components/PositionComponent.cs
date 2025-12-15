using System;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    public class PositionComponent : IElementComponent
    {
        public string ComponentType => "Position";
        public bool IsEnabled { get; set; } = true;
        public ShapeBounds Bounds { get; set; }

        public PositionComponent() { Bounds = new ShapeBounds(); }

        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            try
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) { Bounds = new ShapeBounds(); return; }

                var xfrm = spPr.Transform2D;
                if (xfrm == null) { Bounds = new ShapeBounds(); return; }

                long x = xfrm.Offset?.X ?? 0;
                long y = xfrm.Offset?.Y ?? 0;
                long cx = xfrm.Extents?.Cx ?? 0;
                long cy = xfrm.Extents?.Cy ?? 0;
                float rotation = xfrm.Rotation != null ? (float)(xfrm.Rotation.Value / 60000.0) : 0;

                Bounds = new ShapeBounds();
                Bounds.SetFromEmu(x, y, cx, cy, rotation);
            }
            catch { Bounds = new ShapeBounds(); }
        }

        public void ApplyToShape(Shape shape, SlidePart slidePart)
        {
            if (!IsEnabled || Bounds == null) return;
            var spPr = shape.ShapeProperties;
            if (spPr == null) return;

            var xfrm = spPr.Transform2D;
            if (xfrm == null) { xfrm = new A.Transform2D(); spPr.Transform2D = xfrm; }
            if (xfrm.Offset == null) xfrm.Offset = new A.Offset();
            if (xfrm.Extents == null) xfrm.Extents = new A.Extents();

            xfrm.Offset.X = UnitConverter.CmToEmu(Bounds.X);
            xfrm.Offset.Y = UnitConverter.CmToEmu(Bounds.Y);
            xfrm.Extents.Cx = UnitConverter.CmToEmu(Bounds.Width);
            xfrm.Extents.Cy = UnitConverter.CmToEmu(Bounds.Height);
            if (Bounds.Rotation != 0) xfrm.Rotation = (int)(Bounds.Rotation * 60000);
        }

        public string ToJson()
        {
            if (!IsEnabled || Bounds == null) return "null";
            string boxStr = $"{Bounds.X:F2},{Bounds.Y:F2},{Bounds.Width:F2},{Bounds.Height:F2}";
            return $"\"box\":\"{boxStr}\",\"rotation\":{Bounds.Rotation:F1}";
        }

        public void FromJson(object jsonData)
        {
            try
            {
                if (jsonData == null) return;
                dynamic data = jsonData;
                string box = data.box?.ToString() ?? data["box"]?.ToString();
                double rotation = 0;
                try { if (data.rotation != null) rotation = (double)data.rotation; } catch { }
                if (!string.IsNullOrEmpty(box))
                {
                    var parts = box.Split(',');
                    if (parts.Length == 4)
                    {
                        Bounds = new ShapeBounds {
                            X = float.Parse(parts[0]), Y = float.Parse(parts[1]),
                            Width = float.Parse(parts[2]), Height = float.Parse(parts[3]),
                            Rotation = (float)rotation
                        };
                    }
                }
            }
            catch { Bounds = new ShapeBounds(); }
        }

        public void SetPosition(double x, double y, double width, double height, float rotation = 0)
        {
            Bounds = new ShapeBounds();
            Bounds.SetFromPoints(x, y, width, height, rotation);
        }

        public void SetPositionFromEmu(long x, long y, long cx, long cy, float rotation = 0)
        {
            Bounds = new ShapeBounds();
            Bounds.SetFromEmu(x, y, cx, cy, rotation);
        }
    }
}
