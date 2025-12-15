using System;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// 填充类型枚举 (替代 Office.MsoFillType)
    /// </summary>
    public enum FillType
    {
        NoFill = 0,
        Solid = 1,
        Gradient = 2,
        Pattern = 3,
        Picture = 4,
        Background = 5
    }

    /// <summary>
    /// 填充信息
    /// </summary>
    public class FillInfo
    {
        public bool HasFill { get; set; }
        public ColorInfo Color { get; set; }
        public FillType FillType { get; set; }
        public float Transparency { get; set; }

        public FillInfo()
        {
            HasFill = false;
            Color = new ColorInfo();
            FillType = FillType.NoFill;
            Transparency = 0.0f;
        }

        public FillInfo(bool hasFill, ColorInfo color, FillType fillType)
        {
            HasFill = hasFill;
            Color = color ?? new ColorInfo();
            FillType = fillType;
            Transparency = 0.0f;
        }

        public override string ToString()
        {
            if (!HasFill)
                return "无填充";
            return $"填充类型: {FillType}, 颜色: {Color}";
        }
    }
}
