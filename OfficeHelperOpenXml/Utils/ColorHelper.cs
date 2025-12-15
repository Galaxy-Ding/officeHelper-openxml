using System;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Utils
{
    public static class ColorHelper
    {
        public static string ExtractColorString(A.SolidFill solidFill, SlidePart slidePart = null)
        {
            if (solidFill == null) return "";

            try
            {
                var rgbColor = solidFill.RgbColorModelHex;
                if (rgbColor?.Val != null)
                {
                    return "#" + rgbColor.Val.Value;
                }

                var schemeColor = solidFill.SchemeColor;
                if (schemeColor != null && slidePart != null)
                {
                    var colorInfo = ResolveSchemeColor(schemeColor, slidePart);
                    if (!colorInfo.IsTransparent)
                    {
                        return colorInfo.ToHex();
                    }
                }
                else if (schemeColor?.Val != null)
                {
                    var defaultColor = GetDefaultSchemeColor(schemeColor.Val.Value);
                    if (!defaultColor.IsTransparent)
                    {
                        return defaultColor.ToHex();
                    }
                }
            }
            catch { }

            return "";
        }

        /// <summary>
        /// 从 SolidFill 提取完整的 ColorInfo（包含原始主题色信息）
        /// 用于无损保存和写回
        /// </summary>
        public static ColorInfo ExtractColorInfo(A.SolidFill solidFill, SlidePart slidePart = null)
        {
            if (solidFill == null) return new ColorInfo(0, 0, 0, true);

            try
            {
                // 直接RGB颜色
                var rgbColor = solidFill.RgbColorModelHex;
                if (rgbColor?.Val != null)
                {
                    var colorInfo = ParseHexColor(rgbColor.Val.Value);
                    colorInfo.OriginalHex = rgbColor.Val.Value;
                    colorInfo.IsThemeColor = false;
                    return colorInfo;
                }

                // 主题色
                var schemeColor = solidFill.SchemeColor;
                if (schemeColor != null && slidePart != null)
                {
                    return ResolveSchemeColor(schemeColor, slidePart);
                }
                else if (schemeColor?.Val != null)
                {
                    var colorInfo = GetDefaultSchemeColor(schemeColor.Val.Value);
                    colorInfo.SchemeColorName = schemeColor.Val.Value.ToString();
                    colorInfo.Transforms = ExtractColorTransforms(schemeColor);
                    return colorInfo;
                }
            }
            catch { }

            return new ColorInfo(0, 0, 0, true);
        }

        /// <summary>
        /// 创建用于写入PPT的SolidFill
        /// 优先使用主题色+修改器（无损），否则使用RGB
        /// </summary>
        public static A.SolidFill CreateSolidFill(ColorInfo colorInfo)
        {
            if (colorInfo == null || colorInfo.IsTransparent)
                return null;

            var solidFill = new A.SolidFill();

            // 优先使用主题色（无损）
            if (colorInfo.IsThemeColor && !string.IsNullOrEmpty(colorInfo.SchemeColorName))
            {
                var schemeColor = new A.SchemeColor { Val = ParseSchemeColorValue(colorInfo.SchemeColorName) };

                // 添加原始的颜色修改器
                if (colorInfo.Transforms != null)
                {
                    if (colorInfo.Transforms.LumMod.HasValue)
                        schemeColor.Append(new A.LuminanceModulation { Val = colorInfo.Transforms.LumMod.Value });
                    if (colorInfo.Transforms.LumOff.HasValue)
                        schemeColor.Append(new A.LuminanceOffset { Val = colorInfo.Transforms.LumOff.Value });
                    if (colorInfo.Transforms.Tint.HasValue)
                        schemeColor.Append(new A.Tint { Val = colorInfo.Transforms.Tint.Value });
                    if (colorInfo.Transforms.Shade.HasValue)
                        schemeColor.Append(new A.Shade { Val = colorInfo.Transforms.Shade.Value });
                    if (colorInfo.Transforms.SatMod.HasValue)
                        schemeColor.Append(new A.SaturationModulation { Val = colorInfo.Transforms.SatMod.Value });
                    if (colorInfo.Transforms.Alpha.HasValue)
                        schemeColor.Append(new A.Alpha { Val = colorInfo.Transforms.Alpha.Value });
                }

                solidFill.Append(schemeColor);
            }
            else
            {
                // 使用RGB颜色
                string hexValue = colorInfo.OriginalHex ?? $"{colorInfo.Red:X2}{colorInfo.Green:X2}{colorInfo.Blue:X2}";
                solidFill.Append(new A.RgbColorModelHex { Val = hexValue });
            }

            return solidFill;
        }

        /// <summary>
        /// 将主题色枚举值转换为字符串名称
        /// </summary>
        private static string ConvertSchemeColorToString(A.SchemeColorValues colorVal)
        {
            if (colorVal == A.SchemeColorValues.Dark1) return "dk1";
            if (colorVal == A.SchemeColorValues.Light1) return "lt1";
            if (colorVal == A.SchemeColorValues.Dark2) return "dk2";
            if (colorVal == A.SchemeColorValues.Light2) return "lt2";
            if (colorVal == A.SchemeColorValues.Text1) return "tx1";
            if (colorVal == A.SchemeColorValues.Text2) return "tx2";
            if (colorVal == A.SchemeColorValues.Background1) return "bg1";
            if (colorVal == A.SchemeColorValues.Background2) return "bg2";
            if (colorVal == A.SchemeColorValues.Accent1) return "accent1";
            if (colorVal == A.SchemeColorValues.Accent2) return "accent2";
            if (colorVal == A.SchemeColorValues.Accent3) return "accent3";
            if (colorVal == A.SchemeColorValues.Accent4) return "accent4";
            if (colorVal == A.SchemeColorValues.Accent5) return "accent5";
            if (colorVal == A.SchemeColorValues.Accent6) return "accent6";
            if (colorVal == A.SchemeColorValues.Hyperlink) return "hlink";
            if (colorVal == A.SchemeColorValues.FollowedHyperlink) return "folhlink";
            return "dk1"; // default
        }

        /// <summary>
        /// 解析主题色名称为枚举值
        /// </summary>
        private static A.SchemeColorValues ParseSchemeColorValue(string name)
        {
            switch (name?.ToLower())
            {
                case "dk1":
                case "dark1": return A.SchemeColorValues.Dark1;
                case "lt1":
                case "light1": return A.SchemeColorValues.Light1;
                case "dk2":
                case "dark2": return A.SchemeColorValues.Dark2;
                case "lt2":
                case "light2": return A.SchemeColorValues.Light2;
                case "tx1":
                case "text1": return A.SchemeColorValues.Text1;
                case "tx2":
                case "text2": return A.SchemeColorValues.Text2;
                case "bg1":
                case "background1": return A.SchemeColorValues.Background1;
                case "bg2":
                case "background2": return A.SchemeColorValues.Background2;
                case "accent1": return A.SchemeColorValues.Accent1;
                case "accent2": return A.SchemeColorValues.Accent2;
                case "accent3": return A.SchemeColorValues.Accent3;
                case "accent4": return A.SchemeColorValues.Accent4;
                case "accent5": return A.SchemeColorValues.Accent5;
                case "accent6": return A.SchemeColorValues.Accent6;
                case "hlink":
                case "hyperlink": return A.SchemeColorValues.Hyperlink;
                case "folhlink":
                case "followedhyperlink": return A.SchemeColorValues.FollowedHyperlink;
                default: return A.SchemeColorValues.Dark1;
            }
        }

        public static ColorInfo ParseHexColor(string hexValue)
        {
            if (string.IsNullOrEmpty(hexValue))
                return new ColorInfo(0, 0, 0, true);

            try
            {
                hexValue = hexValue.TrimStart('#');
                if (hexValue.Length == 6)
                {
                    int r = Convert.ToInt32(hexValue.Substring(0, 2), 16);
                    int g = Convert.ToInt32(hexValue.Substring(2, 2), 16);
                    int b = Convert.ToInt32(hexValue.Substring(4, 2), 16);
                    return new ColorInfo(r, g, b, false);
                }
                else if (hexValue.Length == 8)
                {
                    int a = Convert.ToInt32(hexValue.Substring(0, 2), 16);
                    int r = Convert.ToInt32(hexValue.Substring(2, 2), 16);
                    int g = Convert.ToInt32(hexValue.Substring(4, 2), 16);
                    int b = Convert.ToInt32(hexValue.Substring(6, 2), 16);
                    return new ColorInfo(r, g, b, a == 0);
                }
            }
            catch { }
            return new ColorInfo(0, 0, 0, true);
        }

        public static ColorInfo ResolveSchemeColor(A.SchemeColor schemeColor, SlidePart slidePart)
        {
            if (schemeColor == null || schemeColor.Val == null)
                return new ColorInfo(0, 0, 0, true);

            try
            {
                var colorVal = schemeColor.Val.Value;
                var colorInfo = new ColorInfo { IsThemeColor = true };
                int r = 0, g = 0, b = 0;
                bool colorFound = false;

                // 保存原始主题色名称（转换为小写字符串）
                colorInfo.SchemeColorName = ConvertSchemeColorToString(colorVal);

                var themePart = slidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart;
                if (themePart?.Theme?.ThemeElements?.ColorScheme != null)
                {
                    var colorScheme = themePart.Theme.ThemeElements.ColorScheme;
                    A.Color2Type color2Type = GetColor2TypeFromScheme(colorScheme, colorVal, colorInfo);

                    if (color2Type != null)
                    {
                        var rgb = color2Type.RgbColorModelHex;
                        if (rgb != null && rgb.Val != null)
                        {
                            var parsed = ParseHexColor(rgb.Val.Value);
                            r = parsed.Red;
                            g = parsed.Green;
                            b = parsed.Blue;
                            colorFound = true;
                            // 保存原始十六进制值
                            colorInfo.OriginalHex = rgb.Val.Value;
                        }
                        else
                        {
                            var srgb = color2Type.SystemColor;
                            if (srgb != null && srgb.LastColor != null)
                            {
                                var parsed = ParseHexColor(srgb.LastColor.Value);
                                r = parsed.Red;
                                g = parsed.Green;
                                b = parsed.Blue;
                                colorFound = true;
                                colorInfo.OriginalHex = srgb.LastColor.Value;
                            }
                        }
                    }
                }

                if (!colorFound)
                {
                    var defaultColor = GetDefaultSchemeColor(colorVal);
                    r = defaultColor.Red;
                    g = defaultColor.Green;
                    b = defaultColor.Blue;
                }

                // 提取并保存原始颜色修改器参数
                colorInfo.Transforms = ExtractColorTransforms(schemeColor);

                // 应用颜色修改器 (LumMod, LumOff, Tint, Shade 等)
                ApplyColorTransforms(schemeColor, ref r, ref g, ref b);

                colorInfo.Red = r;
                colorInfo.Green = g;
                colorInfo.Blue = b;
                colorInfo.IsTransparent = false;
                return colorInfo;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解析主题颜色时出错: {ex.Message}");
                return new ColorInfo(0, 0, 0, true);
            }
        }

        /// <summary>
        /// 提取颜色修改器参数（不进行计算）
        /// 这些原始参数可用于无损写回PPT
        /// </summary>
        private static ColorTransforms ExtractColorTransforms(A.SchemeColor schemeColor)
        {
            var transforms = new ColorTransforms();

            var lumMod = schemeColor.GetFirstChild<A.LuminanceModulation>();
            var lumOff = schemeColor.GetFirstChild<A.LuminanceOffset>();
            var tint = schemeColor.GetFirstChild<A.Tint>();
            var shade = schemeColor.GetFirstChild<A.Shade>();
            var satMod = schemeColor.GetFirstChild<A.SaturationModulation>();
            var satOff = schemeColor.GetFirstChild<A.SaturationOffset>();
            var alpha = schemeColor.GetFirstChild<A.Alpha>();

            if (lumMod?.Val != null) transforms.LumMod = lumMod.Val.Value;
            if (lumOff?.Val != null) transforms.LumOff = lumOff.Val.Value;
            if (tint?.Val != null) transforms.Tint = tint.Val.Value;
            if (shade?.Val != null) transforms.Shade = shade.Val.Value;
            if (satMod?.Val != null) transforms.SatMod = satMod.Val.Value;
            if (satOff?.Val != null) transforms.SatOff = satOff.Val.Value;
            if (alpha?.Val != null) transforms.Alpha = alpha.Val.Value;

            return transforms.HasTransforms ? transforms : null;
        }

        /// <summary>
        /// 应用颜色变换（亮度调整、色调、阴影等）
        /// 根据 ECMA-376 规范：
        /// - Tint/Shade 直接在 RGB 空间操作
        /// - LumMod/LumOff 在 HSL 空间操作 Luminance
        /// </summary>
        private static void ApplyColorTransforms(A.SchemeColor schemeColor, ref int r, ref int g, ref int b)
        {
            // 先检查是否有任何变换需要应用
            var lumMod = schemeColor.GetFirstChild<A.LuminanceModulation>();
            var lumOff = schemeColor.GetFirstChild<A.LuminanceOffset>();
            var tint = schemeColor.GetFirstChild<A.Tint>();
            var shade = schemeColor.GetFirstChild<A.Shade>();

            // 如果没有任何变换，直接返回
            if (lumMod == null && lumOff == null && tint == null && shade == null)
            {
                return;
            }

            // Tint - 向白色混合 (直接在 RGB 空间操作)
            // 公式: newColor = color + (255 - color) * (1 - tint/100000)
            // 使用向下取整以匹配 PowerPoint 的行为
            if (tint?.Val != null)
            {
                double tintValue = tint.Val.Value / 100000.0;
                double factor = 1.0 - tintValue;
                r = (int)Math.Floor(r + (255.0 - r) * factor);
                g = (int)Math.Floor(g + (255.0 - g) * factor);
                b = (int)Math.Floor(b + (255.0 - b) * factor);
            }

            // Shade - 向黑色混合 (直接在 RGB 空间操作)
            // 公式: newColor = color * (shade/100000)
            // 使用向下取整以匹配 PowerPoint 的行为
            if (shade?.Val != null)
            {
                double shadeValue = shade.Val.Value / 100000.0;
                r = (int)Math.Floor(r * shadeValue);
                g = (int)Math.Floor(g * shadeValue);
                b = (int)Math.Floor(b * shadeValue);
            }

            // LumMod/LumOff - 在 HSL 空间操作 Luminance
            if (lumMod != null || lumOff != null)
            {
                double h, s, l;
                RgbToHsl(r, g, b, out h, out s, out l);

                // LumMod - 亮度调制 (百分比，如 95000 = 95%)
                if (lumMod?.Val != null)
                {
                    double modValue = lumMod.Val.Value / 100000.0;
                    l = l * modValue;
                }

                // LumOff - 亮度偏移 (百分比，如 5000 = 5%)
                if (lumOff?.Val != null)
                {
                    double offValue = lumOff.Val.Value / 100000.0;
                    l = l + offValue;
                }

                // 限制范围
                l = Math.Max(0, Math.Min(1, l));

                // 转回RGB
                HslToRgb(h, s, l, out r, out g, out b);
            }

            // 确保值在有效范围内
            r = Math.Max(0, Math.Min(255, r));
            g = Math.Max(0, Math.Min(255, g));
            b = Math.Max(0, Math.Min(255, b));
        }

        /// <summary>
        /// RGB转HSL
        /// </summary>
        private static void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
        {
            double rd = r / 255.0;
            double gd = g / 255.0;
            double bd = b / 255.0;

            double max = Math.Max(rd, Math.Max(gd, bd));
            double min = Math.Min(rd, Math.Min(gd, bd));

            h = s = l = (max + min) / 2.0;

            if (max == min)
            {
                h = s = 0;
            }
            else
            {
                double d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

                if (max == rd)
                    h = (gd - bd) / d + (gd < bd ? 6 : 0);
                else if (max == gd)
                    h = (bd - rd) / d + 2;
                else
                    h = (rd - gd) / d + 4;

                h /= 6;
            }
        }

        /// <summary>
        /// HSL转RGB
        /// 使用标准四舍五入 (MidpointRounding.AwayFromZero) 以匹配 PowerPoint 的行为
        /// </summary>
        private static void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
        {
            double rd, gd, bd;

            if (s == 0)
            {
                rd = gd = bd = l;
            }
            else
            {
                double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
                double p = 2 * l - q;
                rd = HueToRgb(p, q, h + 1.0 / 3);
                gd = HueToRgb(p, q, h);
                bd = HueToRgb(p, q, h - 1.0 / 3);
            }

            // 使用自定义舍入策略以匹配 PowerPoint 的行为
            // 对于恰好是 0.5 的情况（如 127.5），向下取整到 127
            // 对于其他情况，使用标准四舍五入
            r = RoundColorComponent(rd * 255);
            g = RoundColorComponent(gd * 255);
            b = RoundColorComponent(bd * 255);
        }

        /// <summary>
        /// 自定义颜色分量舍入策略
        /// 对于恰好是 0.5 的情况向下取整，其他情况使用标准四舍五入
        /// 这样可以匹配 PowerPoint 对于 lumMod=50000 的行为
        /// </summary>
        private static int RoundColorComponent(double value)
        {
            double fractional = value - Math.Floor(value);
            // 如果小数部分恰好是 0.5，向下取整
            if (Math.Abs(fractional - 0.5) < 0.0001)
            {
                return (int)Math.Floor(value);
            }
            // 否则使用标准四舍五入
            return (int)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }

        private static A.Color2Type GetColor2TypeFromScheme(A.ColorScheme colorScheme, A.SchemeColorValues colorVal, ColorInfo colorInfo)
        {
            // 处理 Dark1/Light1/Dark2/Light2 及其别名 (Text1/Background1/Text2/Background2)
            // bg1 (Background1) -> Light1, tx1 (Text1) -> Dark1
            // bg2 (Background2) -> Light2, tx2 (Text2) -> Dark2
            if (colorVal == A.SchemeColorValues.Dark1 || colorVal == A.SchemeColorValues.Text1)
            {
                colorInfo.SchemeColorIndex = 1;
                return colorScheme.Dark1Color;
            }
            if (colorVal == A.SchemeColorValues.Light1 || colorVal == A.SchemeColorValues.Background1)
            {
                colorInfo.SchemeColorIndex = 2;
                return colorScheme.Light1Color;
            }
            if (colorVal == A.SchemeColorValues.Dark2 || colorVal == A.SchemeColorValues.Text2)
            {
                colorInfo.SchemeColorIndex = 3;
                return colorScheme.Dark2Color;
            }
            if (colorVal == A.SchemeColorValues.Light2 || colorVal == A.SchemeColorValues.Background2)
            {
                colorInfo.SchemeColorIndex = 4;
                return colorScheme.Light2Color;
            }
            if (colorVal == A.SchemeColorValues.Accent1) { colorInfo.SchemeColorIndex = 5; return colorScheme.Accent1Color; }
            if (colorVal == A.SchemeColorValues.Accent2) { colorInfo.SchemeColorIndex = 6; return colorScheme.Accent2Color; }
            if (colorVal == A.SchemeColorValues.Accent3) { colorInfo.SchemeColorIndex = 7; return colorScheme.Accent3Color; }
            if (colorVal == A.SchemeColorValues.Accent4) { colorInfo.SchemeColorIndex = 8; return colorScheme.Accent4Color; }
            if (colorVal == A.SchemeColorValues.Accent5) { colorInfo.SchemeColorIndex = 9; return colorScheme.Accent5Color; }
            if (colorVal == A.SchemeColorValues.Accent6) { colorInfo.SchemeColorIndex = 10; return colorScheme.Accent6Color; }
            if (colorVal == A.SchemeColorValues.Hyperlink) { colorInfo.SchemeColorIndex = 11; return colorScheme.Hyperlink; }
            if (colorVal == A.SchemeColorValues.FollowedHyperlink) { colorInfo.SchemeColorIndex = 12; return colorScheme.FollowedHyperlinkColor; }
            return null;
        }

        private static ColorInfo GetDefaultSchemeColor(A.SchemeColorValues colorVal)
        {
            // 处理 Dark1/Light1/Dark2/Light2 及其别名 (Text1/Background1/Text2/Background2)
            if (colorVal == A.SchemeColorValues.Dark1 || colorVal == A.SchemeColorValues.Text1)
                return new ColorInfo(0, 0, 0, false) { IsThemeColor = true, SchemeColorIndex = 1 };
            if (colorVal == A.SchemeColorValues.Light1 || colorVal == A.SchemeColorValues.Background1)
                return new ColorInfo(255, 255, 255, false) { IsThemeColor = true, SchemeColorIndex = 2 };
            if (colorVal == A.SchemeColorValues.Dark2 || colorVal == A.SchemeColorValues.Text2)
                return new ColorInfo(68, 84, 106, false) { IsThemeColor = true, SchemeColorIndex = 3 };
            if (colorVal == A.SchemeColorValues.Light2 || colorVal == A.SchemeColorValues.Background2)
                return new ColorInfo(231, 230, 230, false) { IsThemeColor = true, SchemeColorIndex = 4 };
            if (colorVal == A.SchemeColorValues.Accent1) return new ColorInfo(68, 114, 196, false) { IsThemeColor = true, SchemeColorIndex = 5 };
            if (colorVal == A.SchemeColorValues.Accent2) return new ColorInfo(237, 125, 49, false) { IsThemeColor = true, SchemeColorIndex = 6 };
            if (colorVal == A.SchemeColorValues.Accent3) return new ColorInfo(165, 165, 165, false) { IsThemeColor = true, SchemeColorIndex = 7 };
            if (colorVal == A.SchemeColorValues.Accent4) return new ColorInfo(255, 192, 0, false) { IsThemeColor = true, SchemeColorIndex = 8 };
            if (colorVal == A.SchemeColorValues.Accent5) return new ColorInfo(91, 155, 213, false) { IsThemeColor = true, SchemeColorIndex = 9 };
            if (colorVal == A.SchemeColorValues.Accent6) return new ColorInfo(112, 173, 71, false) { IsThemeColor = true, SchemeColorIndex = 10 };
            if (colorVal == A.SchemeColorValues.Hyperlink) return new ColorInfo(5, 99, 193, false) { IsThemeColor = true, SchemeColorIndex = 11 };
            if (colorVal == A.SchemeColorValues.FollowedHyperlink) return new ColorInfo(149, 79, 114, false) { IsThemeColor = true, SchemeColorIndex = 12 };
            return new ColorInfo(0, 0, 0, true);
        }

        // ========== JSON-to-PPTX Conversion Methods ==========

        /// <summary>
        /// Parse RGB color string from JSON format "RGB(r,g,b)" to hex string.
        /// Handles edge cases like whitespace and invalid values.
        /// </summary>
        /// <param name="rgbString">RGB string in format "RGB(r,g,b)" where r, g, b are 0-255</param>
        /// <returns>Hex color string without # prefix (e.g., "FF0000" for red), or empty string if invalid</returns>
        /// <example>
        /// <code>
        /// string hex = ColorHelper.ParseRgbToHex("RGB(255, 0, 0)");  // Returns "FF0000"
        /// string hex2 = ColorHelper.ParseRgbToHex("RGB(128,128,128)"); // Returns "808080"
        /// </code>
        /// </example>
        /// <remarks>
        /// Validates that all RGB components are in the range 0-255.
        /// Trims whitespace from input and component values.
        /// Returns empty string for malformed input.
        /// </remarks>
        public static string ParseRgbToHex(string rgbString)
        {
            if (string.IsNullOrWhiteSpace(rgbString))
                return "";

            try
            {
                // Remove whitespace and convert to uppercase
                rgbString = rgbString.Trim().ToUpper();

                // Check if it starts with "RGB("
                if (!rgbString.StartsWith("RGB(") || !rgbString.EndsWith(")"))
                    return "";

                // Extract the content between parentheses
                string content = rgbString.Substring(4, rgbString.Length - 5);

                // Split by comma
                string[] parts = content.Split(',');
                if (parts.Length != 3)
                    return "";

                // Parse each component
                int r = int.Parse(parts[0].Trim());
                int g = int.Parse(parts[1].Trim());
                int b = int.Parse(parts[2].Trim());

                // Validate range
                if (r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255)
                    return "";

                // Convert to hex
                return $"{r:X2}{g:X2}{b:X2}";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Map schemeColor name from JSON to OpenXML SchemeColorValues.
        /// Supports all theme colors (bg1, bg2, tx1, tx2, accent1-6, hyperlink, etc.).
        /// </summary>
        /// <param name="schemeColorName">Scheme color name from JSON (e.g., "bg1", "accent1", "tx1")</param>
        /// <returns>OpenXML SchemeColorValues enum value, defaults to Dark1 if name is invalid</returns>
        /// <remarks>
        /// Supported color names:
        /// - Background: bg1, bg2 (or background1, background2)
        /// - Text: tx1, tx2 (or text1, text2)
        /// - Dark/Light: dk1, dk2, lt1, lt2
        /// - Accents: accent1 through accent6
        /// - Links: hlink, folhlink (hyperlink, followedhyperlink)
        /// Case-insensitive matching.
        /// </remarks>
        public static A.SchemeColorValues MapSchemeColorName(string schemeColorName)
        {
            if (string.IsNullOrWhiteSpace(schemeColorName))
                return A.SchemeColorValues.Dark1; // Default

            // Use existing ParseSchemeColorValue method which already handles all mappings
            return ParseSchemeColorValue(schemeColorName);
        }

        /// <summary>
        /// Create OpenXML SchemeColor with color transforms applied.
        /// Applies lumMod (luminance modulation) and lumOff (luminance offset) to theme colors.
        /// </summary>
        /// <param name="schemeColorName">Scheme color name from JSON (e.g., "accent1")</param>
        /// <param name="lumMod">Luminance modulation value in percentage * 1000 (e.g., 60000 = 60%)</param>
        /// <param name="lumOff">Luminance offset value in percentage * 1000 (e.g., 40000 = 40%)</param>
        /// <returns>OpenXML SchemeColor with transforms applied</returns>
        /// <example>
        /// <code>
        /// // Create accent1 color with 60% luminance and 40% offset
        /// var color = ColorHelper.CreateSchemeColorWithTransforms("accent1", 60000, 40000);
        /// </code>
        /// </example>
        /// <remarks>
        /// Transform values are in the range 0-100000 (representing 0-100%).
        /// LumMod multiplies the luminance by the specified percentage.
        /// LumOff adds the specified percentage to the luminance.
        /// Null values for lumMod or lumOff mean no transform is applied.
        /// </remarks>
        public static A.SchemeColor CreateSchemeColorWithTransforms(string schemeColorName, int? lumMod, int? lumOff)
        {
            var schemeColor = new A.SchemeColor
            {
                Val = MapSchemeColorName(schemeColorName)
            };

            // Apply luminance modulation if specified
            if (lumMod.HasValue)
            {
                schemeColor.Append(new A.LuminanceModulation { Val = lumMod.Value });
            }

            // Apply luminance offset if specified
            if (lumOff.HasValue)
            {
                schemeColor.Append(new A.LuminanceOffset { Val = lumOff.Value });
            }

            return schemeColor;
        }

        /// <summary>
        /// Create appropriate OpenXML color object from JSON data.
        /// Prioritizes schemeColor over RGB when both are present, as per design specification.
        /// </summary>
        /// <param name="rgbColor">RGB color string in format "RGB(r,g,b)"</param>
        /// <param name="schemeColor">Scheme color name (e.g., "bg1", "accent1")</param>
        /// <param name="lumMod">Optional luminance modulation (0-100000)</param>
        /// <param name="lumOff">Optional luminance offset (0-100000)</param>
        /// <returns>OpenXML color object (SchemeColor or RgbColorModelHex), or null if no valid color</returns>
        /// <example>
        /// <code>
        /// // Create theme color with transforms
        /// var color1 = ColorHelper.CreateColorFromJson(null, "accent1", 60000, 40000);
        /// 
        /// // Create RGB color
        /// var color2 = ColorHelper.CreateColorFromJson("RGB(255,0,0)", null, null, null);
        /// 
        /// // Theme color takes priority
        /// var color3 = ColorHelper.CreateColorFromJson("RGB(255,0,0)", "accent1", null, null);
        /// // Returns SchemeColor for accent1, not RGB
        /// </code>
        /// </example>
        /// <remarks>
        /// Priority order:
        /// 1. SchemeColor with transforms (if schemeColor is specified)
        /// 2. RGB color (if rgbColor is specified and schemeColor is not)
        /// 3. null (if neither is specified or both are invalid)
        /// This ensures theme colors are preserved for better PowerPoint compatibility.
        /// </remarks>
        public static object CreateColorFromJson(string rgbColor, string schemeColor, int? lumMod, int? lumOff)
        {
            // Prioritize schemeColor over RGB
            if (!string.IsNullOrWhiteSpace(schemeColor))
            {
                return CreateSchemeColorWithTransforms(schemeColor, lumMod, lumOff);
            }

            // Fall back to RGB color
            if (!string.IsNullOrWhiteSpace(rgbColor))
            {
                string hexValue = ParseRgbToHex(rgbColor);
                if (!string.IsNullOrEmpty(hexValue))
                {
                    return new A.RgbColorModelHex { Val = hexValue };
                }
            }

            // Return null if no valid color
            return null;
        }
    }
}
