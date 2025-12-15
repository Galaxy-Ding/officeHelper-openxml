using System;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// 线条虚线样式枚举 (替代 Office.MsoLineDashStyle)
    /// </summary>
    public enum LineDashStyle
    {
        Solid = 1,
        SquareDot = 2,
        RoundDot = 3,
        Dash = 4,
        DashDot = 5,
        LongDash = 6,
        LongDashDot = 7,
        LongDashDotDot = 8
    }

    /// <summary>
    /// 箭头样式枚举 (替代 Office.MsoArrowheadStyle)
    /// </summary>
    public enum ArrowheadStyle
    {
        None = 1,
        Triangle = 2,
        Stealth = 3,
        Diamond = 4,
        Oval = 5,
        Open = 6
    }

    /// <summary>
    /// 箭头长度枚举
    /// </summary>
    public enum ArrowheadLength
    {
        Short = 1,
        Medium = 2,
        Long = 3
    }

    /// <summary>
    /// 箭头宽度枚举
    /// </summary>
    public enum ArrowheadWidth
    {
        Narrow = 1,
        Medium = 2,
        Wide = 3
    }

    /// <summary>
    /// 线条样式信息
    /// </summary>
    public class LineStyle
    {
        public ColorInfo Color { get; set; }
        public bool HasOutline { get; set; }
        public float Weight { get; set; }
        public float Transparency { get; set; }
        public LineDashStyle DashStyle { get; set; }
        public string DashStyleName { get; set; }
        public bool HasBeginArrow { get; set; }
        public ArrowheadStyle BeginArrowStyle { get; set; }
        public ArrowheadLength BeginArrowLength { get; set; }
        public ArrowheadWidth BeginArrowWidth { get; set; }
        public bool HasEndArrow { get; set; }
        public ArrowheadStyle EndArrowStyle { get; set; }
        public ArrowheadLength EndArrowLength { get; set; }
        public ArrowheadWidth EndArrowWidth { get; set; }

        public LineStyle()
        {
            Color = new ColorInfo();
            HasOutline = false;
            Weight = 0;
            Transparency = 0.0f;
            DashStyle = LineDashStyle.Solid;
            DashStyleName = "无边框";
            HasBeginArrow = false;
            BeginArrowStyle = ArrowheadStyle.None;
            BeginArrowLength = ArrowheadLength.Medium;
            BeginArrowWidth = ArrowheadWidth.Medium;
            HasEndArrow = false;
            EndArrowStyle = ArrowheadStyle.None;
            EndArrowLength = ArrowheadLength.Medium;
            EndArrowWidth = ArrowheadWidth.Medium;
        }

        public LineStyle(ColorInfo color, float weight, LineDashStyle dashStyle, string dashStyleName)
        {
            Color = color ?? new ColorInfo();
            HasOutline = weight > 0;
            Weight = weight;
            Transparency = 0.0f;
            DashStyle = dashStyle;
            DashStyleName = dashStyleName ?? "无边框";
            HasBeginArrow = false;
            BeginArrowStyle = ArrowheadStyle.None;
            BeginArrowLength = ArrowheadLength.Medium;
            BeginArrowWidth = ArrowheadWidth.Medium;
            HasEndArrow = false;
            EndArrowStyle = ArrowheadStyle.None;
            EndArrowLength = ArrowheadLength.Medium;
            EndArrowWidth = ArrowheadWidth.Medium;
        }

        public override string ToString()
        {
            return $"颜色: {Color}, 粗细: {Weight:F2}pt, 样式: {DashStyleName}";
        }
    }
}
