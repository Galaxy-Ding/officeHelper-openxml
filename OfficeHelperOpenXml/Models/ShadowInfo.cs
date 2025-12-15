using System;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// 阴影类型枚举
    /// </summary>
    public enum ShadowType
    {
        None = 0,
        Outer = 1,
        Inner = 2,
        Perspective = 3
    }

    /// <summary>
    /// 阴影样式枚举
    /// </summary>
    public enum ShadowStyle
    {
        None = 0,
        Preset = 1,
        Custom = 2
    }

    /// <summary>
    /// 阴影信息
    /// </summary>
    public class ShadowInfo
    {
        public bool HasShadow { get; set; }
        public ShadowType Type { get; set; }
        public ShadowStyle Style { get; set; }
        public string ShadowTypeName { get; set; }
        public ColorInfo Color { get; set; }
        public float Blur { get; set; }
        public float Distance { get; set; }
        public float Angle { get; set; }
        public float Transparency { get; set; }
        public float Opacity { get; set; }
        public float OffsetX { get; set; }
        public float OffsetY { get; set; }
        public float Size { get; set; }

        public ShadowInfo()
        {
            HasShadow = false;
            Type = ShadowType.None;
            Style = ShadowStyle.None;
            ShadowTypeName = "";
            Color = new ColorInfo(0, 0, 0, false);
            Blur = 0;
            Distance = 0;
            Angle = 0;
            Transparency = 0;
            Opacity = 0;
            OffsetX = 0;
            OffsetY = 0;
            Size = 0;
        }

        public override string ToString()
        {
            if (!HasShadow)
                return "无阴影";
            return $"阴影类型: {Type}, 颜色: {Color}, 模糊: {Blur}, 距离: {Distance}";
        }
    }
}
