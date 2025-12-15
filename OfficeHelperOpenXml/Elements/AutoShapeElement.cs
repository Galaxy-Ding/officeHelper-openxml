using System;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Components;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Elements
{
    public class AutoShapeElement : BaseElement
    {
        public override string ElementType => "AutoShape";
        
        public string ShapeType { get; set; }
        public int ShapeTypeCode { get; set; }
        
        public AutoShapeElement() : base()
        {
            ShapeType = "Rectangle";
            ShapeTypeCode = 1;
        }
        
        protected override void InitializeComponents(Shape shape, SlidePart slidePart)
        {
            ExtractShapeType(shape);
            
            var positionComponent = new PositionComponent();
            positionComponent.ExtractFromShape(shape, slidePart);
            AddComponent(positionComponent);
            
            var fillComponent = new FillComponent();
            fillComponent.ExtractFromShape(shape, slidePart);
            AddComponent(fillComponent);
            
            var lineComponent = new LineComponent();
            lineComponent.ExtractFromShape(shape, slidePart);
            AddComponent(lineComponent);
            
            var shadowComponent = new ShadowComponent();
            shadowComponent.ExtractFromShape(shape, slidePart);
            AddComponent(shadowComponent);
            
            var textComponent = new TextComponent();
            textComponent.ExtractFromShape(shape, slidePart);
            AddComponent(textComponent);
        }
        
        private void ExtractShapeType(Shape shape)
        {
            try
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) return;
                
                var prstGeom = spPr.GetFirstChild<A.PresetGeometry>();
                if (prstGeom != null && prstGeom.Preset != null)
                {
                    var preset = prstGeom.Preset.Value;
                    ShapeType = GetShapeTypeName(preset);
                    ShapeTypeCode = GetShapeTypeCode(preset);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取形状类型时出错: {ex.Message}");
            }
        }
        
        private int GetShapeTypeCode(A.ShapeTypeValues shapeType)
        {
            if (shapeType == A.ShapeTypeValues.Rectangle) return 1;
            if (shapeType == A.ShapeTypeValues.RoundRectangle) return 5;
            if (shapeType == A.ShapeTypeValues.Ellipse) return 9;
            if (shapeType == A.ShapeTypeValues.Triangle) return 7;
            if (shapeType == A.ShapeTypeValues.RightTriangle) return 6;
            if (shapeType == A.ShapeTypeValues.Parallelogram) return 8;
            if (shapeType == A.ShapeTypeValues.Trapezoid) return 10;
            if (shapeType == A.ShapeTypeValues.Diamond) return 4;
            if (shapeType == A.ShapeTypeValues.Pentagon) return 11;
            if (shapeType == A.ShapeTypeValues.Hexagon) return 12;
            if (shapeType == A.ShapeTypeValues.Line) return 20;
            return 1;
        }
        
        private string GetShapeTypeName(A.ShapeTypeValues shapeType)
        {
            if (shapeType == A.ShapeTypeValues.Rectangle) return "矩形";
            if (shapeType == A.ShapeTypeValues.RoundRectangle) return "圆角矩形";
            if (shapeType == A.ShapeTypeValues.Ellipse) return "椭圆";
            if (shapeType == A.ShapeTypeValues.Triangle) return "三角形";
            if (shapeType == A.ShapeTypeValues.RightTriangle) return "直角三角形";
            if (shapeType == A.ShapeTypeValues.Parallelogram) return "平行四边形";
            if (shapeType == A.ShapeTypeValues.Trapezoid) return "梯形";
            if (shapeType == A.ShapeTypeValues.Diamond) return "菱形";
            if (shapeType == A.ShapeTypeValues.Pentagon) return "五边形";
            if (shapeType == A.ShapeTypeValues.Hexagon) return "六边形";
            if (shapeType == A.ShapeTypeValues.Line) return "直线";
            if (shapeType == A.ShapeTypeValues.RightArrow) return "右箭头";
            if (shapeType == A.ShapeTypeValues.LeftArrow) return "左箭头";
            if (shapeType == A.ShapeTypeValues.Heart) return "心形";
            if (shapeType == A.ShapeTypeValues.Cloud) return "云形";
            return shapeType.ToString();
        }
        
        public override string ToJson()
        {
            // Use base implementation which now matches template format
            return base.ToJson();
        }
    }
}
