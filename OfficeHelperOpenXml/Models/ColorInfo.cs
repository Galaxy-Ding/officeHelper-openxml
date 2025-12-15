using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// 颜色变换信息 - 保存原始的颜色修改器参数
    /// 用于无损写回PPT
    /// </summary>
    public class ColorTransforms
    {
        /// <summary>亮度调制 (Luminance Modulation), 如 50000 = 50%</summary>
        public int? LumMod { get; set; }
        /// <summary>亮度偏移 (Luminance Offset), 如 50000 = 50%</summary>
        public int? LumOff { get; set; }
        /// <summary>色调 (Tint), 如 50000 = 50%, 向白色混合</summary>
        public int? Tint { get; set; }
        /// <summary>阴影 (Shade), 如 50000 = 50%, 向黑色混合</summary>
        public int? Shade { get; set; }
        /// <summary>饱和度调制 (Saturation Modulation)</summary>
        public int? SatMod { get; set; }
        /// <summary>饱和度偏移 (Saturation Offset)</summary>
        public int? SatOff { get; set; }
        /// <summary>透明度 (Alpha), 如 50000 = 50%</summary>
        public int? Alpha { get; set; }

        public bool HasTransforms => LumMod.HasValue || LumOff.HasValue ||
                                      Tint.HasValue || Shade.HasValue ||
                                      SatMod.HasValue || SatOff.HasValue ||
                                      Alpha.HasValue;
    }

    /// <summary>
    /// 颜色信息
    /// </summary>
    public class ColorInfo
    {
        // ===== 计算后的RGB值（用于预览和显示）=====
        public int Red { get; set; }
        public int Green { get; set; }
        public int Blue { get; set; }
        public bool IsTransparent { get; set; }

        // ===== 原始主题色信息（用于无损写回PPT）=====
        public bool IsThemeColor { get; set; }
        /// <summary>主题色索引 (1=dk1, 2=lt1, 3=dk2, 4=lt2, 5-10=accent1-6, 11=hlink, 12=folHlink)</summary>
        public int? SchemeColorIndex { get; set; }
        /// <summary>主题色名称 (如 "accent1", "dk1", "tx1" 等)</summary>
        public string SchemeColorName { get; set; }
        /// <summary>原始颜色修改器参数</summary>
        public ColorTransforms Transforms { get; set; }
        /// <summary>原始RGB十六进制值 (如 "5B9BD5")</summary>
        public string OriginalHex { get; set; }

        public ColorInfo() { Red = 0; Green = 0; Blue = 0; IsTransparent = true; IsThemeColor = false; }
        public ColorInfo(int red, int green, int blue, bool isTransparent = false)
        {
            Red = red; Green = green; Blue = blue; IsTransparent = isTransparent; IsThemeColor = false;
        }

        public override string ToString() => $"RGB({Red}, {Green}, {Blue})";
        public string ToHex() => IsTransparent ? "#00000000" : $"#{Red:X2}{Green:X2}{Blue:X2}";
        public string ToRgbString() => IsTransparent ? "transparent" : $"rgb({Red}, {Green}, {Blue})";

        /// <summary>
        /// 获取用于写入PPT的最佳格式
        /// 优先返回原始格式信息，避免转换误差
        /// </summary>
        public string GetWriteFormat()
        {
            // 优先使用主题色
            if (IsThemeColor && !string.IsNullOrEmpty(SchemeColorName))
            {
                if (Transforms?.HasTransforms == true)
                    return $"scheme:{SchemeColorName}+transforms";
                return $"scheme:{SchemeColorName}";
            }
            // 其次使用原始十六进制
            if (!string.IsNullOrEmpty(OriginalHex))
                return $"hex:{OriginalHex}";
            // 最后使用计算后的RGB
            return $"rgb:{Red:X2}{Green:X2}{Blue:X2}";
        }

        public static ColorInfo Parse(string colorString)
        {
            if (string.IsNullOrWhiteSpace(colorString)) return new ColorInfo();
            colorString = colorString.Trim();
            if (colorString.Equals("transparent", StringComparison.OrdinalIgnoreCase))
                return new ColorInfo { IsTransparent = true };
            if (colorString.StartsWith("RGB(", StringComparison.OrdinalIgnoreCase) ||
                colorString.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    int startIdx = colorString.IndexOf('(') + 1;
                    int endIdx = colorString.IndexOf(')');
                    if (endIdx > startIdx)
                    {
                        string rgbValues = colorString.Substring(startIdx, endIdx - startIdx);
                        string[] parts = rgbValues.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 3)
                        {
                            int r = int.Parse(parts[0].Trim());
                            int g = int.Parse(parts[1].Trim());
                            int b = int.Parse(parts[2].Trim());
                            return new ColorInfo(r, g, b, false);
                        }
                    }
                }
                catch { return new ColorInfo(); }
            }
            if (colorString.StartsWith("#"))
            {
                try
                {
                    string hex = colorString.Substring(1);
                    if (hex.Length == 6)
                    {
                        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                        return new ColorInfo(r, g, b, false);
                    }
                    else if (hex.Length == 8)
                    {
                        int a = Convert.ToInt32(hex.Substring(0, 2), 16);
                        int r = Convert.ToInt32(hex.Substring(2, 2), 16);
                        int g = Convert.ToInt32(hex.Substring(4, 2), 16);
                        int b = Convert.ToInt32(hex.Substring(6, 2), 16);
                        return new ColorInfo(r, g, b, a == 0);
                    }
                }
                catch { return new ColorInfo(); }
            }
            return new ColorInfo();
        }
    }
}
