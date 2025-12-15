using OfficeHelperOpenXml.Utils;

namespace OfficeHelperOpenXml.Models
{
    public class ShapeBounds
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public float Rotation { get; set; }
        
        // 别名属性
        public float X { get => Left; set => Left = value; }
        public float Y { get => Top; set => Top = value; }

        public ShapeBounds() { }

        public ShapeBounds(float left, float top, float width, float height, float rotation = 0)
        {
            Left = left; Top = top; Width = width; Height = height; Rotation = rotation;
        }

        public void SetFromPoints(double x, double y, double width, double height, float rotation = 0)
        {
            Left = (float)UnitConverter.PointsToCm(x);
            Top = (float)UnitConverter.PointsToCm(y);
            Width = (float)UnitConverter.PointsToCm(width);
            Height = (float)UnitConverter.PointsToCm(height);
            Rotation = rotation;
        }

        public void SetFromEmu(long x, long y, long width, long height, float rotation = 0)
        {
            Left = (float)UnitConverter.EmuToCm(x);
            Top = (float)UnitConverter.EmuToCm(y);
            Width = (float)UnitConverter.EmuToCm(width);
            Height = (float)UnitConverter.EmuToCm(height);
            Rotation = rotation;
        }

        public override string ToString()
        {
            return $"位置: ({Left:F2}cm, {Top:F2}cm), 尺寸: {Width:F2}cm x {Height:F2}cm, 旋转: {Rotation:F2}°";
        }
    }
}
