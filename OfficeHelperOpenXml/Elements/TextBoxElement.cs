using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Components;

namespace OfficeHelperOpenXml.Elements
{
    public class TextBoxElement : BaseElement
    {
        public override string ElementType => "TextBox";
        
        public TextBoxElement() : base() { }
        
        protected override void InitializeComponents(Shape shape, SlidePart slidePart)
        {
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
    }
}
