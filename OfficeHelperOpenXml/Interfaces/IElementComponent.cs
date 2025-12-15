using System;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeHelperOpenXml.Interfaces
{
    /// <summary>
    /// 元素组件接口 - 所有组件的基础接口
    /// </summary>
    public interface IElementComponent
    {
        string ComponentType { get; }
        bool IsEnabled { get; set; }
        void ExtractFromShape(Shape shape, SlidePart slidePart);
        void ApplyToShape(Shape shape, SlidePart slidePart);
        string ToJson();
        void FromJson(object jsonData);
    }
}
