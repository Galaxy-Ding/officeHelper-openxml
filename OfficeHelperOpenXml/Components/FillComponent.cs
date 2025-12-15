using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    /// <summary>
    /// 填充组件 - 使用 OpenXML 提取填充信息
    /// </summary>
    public class FillComponent : IElementComponent
    {
        public string ComponentType => "Fill";
        public bool IsEnabled { get; set; } = true;
        
        public FillInfo Fill { get; set; }
        
        public FillComponent()
        {
            Fill = new FillInfo();
        }
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            try
            {
                Fill = new FillInfo();
                
                // 获取形状属性
                var spPr = shape.ShapeProperties;
                if (spPr == null)
                {
                    Fill.HasFill = false;
                    return;
                }
                
                // 检查是否有 NoFill
                var noFill = spPr.GetFirstChild<A.NoFill>();
                if (noFill != null)
                {
                    Fill.HasFill = false;
                    Fill.FillType = FillType.NoFill;
                    return;
                }
                
                // 检查实心填充
                var solidFill = spPr.GetFirstChild<A.SolidFill>();
                if (solidFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Solid;
                    Fill.Color = ExtractColorFromSolidFill(solidFill, slidePart);
                    return;
                }
                
                // 检查渐变填充
                var gradFill = spPr.GetFirstChild<A.GradientFill>();
                if (gradFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Gradient;
                    // 简化处理：取渐变的第一个颜色
                    var gradStops = gradFill.GradientStopList;
                    if (gradStops != null)
                    {
                        var firstStop = gradStops.GetFirstChild<A.GradientStop>();
                        if (firstStop != null)
                        {
                            Fill.Color = ExtractColorFromGradientStop(firstStop, slidePart);
                        }
                    }
                    return;
                }
                
                // 检查图案填充
                var pattFill = spPr.GetFirstChild<A.PatternFill>();
                if (pattFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Pattern;
                    var fgClr = pattFill.ForegroundColor;
                    if (fgClr != null)
                    {
                        Fill.Color = ExtractColorFromColorType(fgClr, slidePart);
                    }
                    return;
                }
                
                // 检查图片填充
                var blipFill = spPr.GetFirstChild<A.BlipFill>();
                if (blipFill != null)
                {
                    Fill.HasFill = true;
                    Fill.FillType = FillType.Picture;
                    return;
                }
                
                // 默认无填充
                Fill.HasFill = false;
                Fill.FillType = FillType.NoFill;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取填充信息时出错: {ex.Message}");
                Fill.HasFill = false;
                Fill.Color = new ColorInfo(0, 0, 0, true);
            }
        }
        
        private ColorInfo ExtractColorFromSolidFill(A.SolidFill solidFill, SlidePart slidePart)
        {
            // RGB颜色
            var rgbColor = solidFill.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null)
            {
                return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            }
            
            // sRGB颜色
            var srgbColor = solidFill.RgbColorModelPercentage;
            if (srgbColor != null)
            {
                int r = (int)(srgbColor.RedPortion?.Value ?? 0) * 255 / 100000;
                int g = (int)(srgbColor.GreenPortion?.Value ?? 0) * 255 / 100000;
                int b = (int)(srgbColor.BluePortion?.Value ?? 0) * 255 / 100000;
                return new ColorInfo(r, g, b, false);
            }
            
            // 主题颜色
            var schemeColor = solidFill.SchemeColor;
            if (schemeColor != null)
            {
                return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            }
            
            return new ColorInfo(0, 0, 0, true);
        }
        
        private ColorInfo ExtractColorFromGradientStop(A.GradientStop stop, SlidePart slidePart)
        {
            var rgbColor = stop.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null)
            {
                return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            }
            
            var schemeColor = stop.SchemeColor;
            if (schemeColor != null)
            {
                return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            }
            
            return new ColorInfo(0, 0, 0, true);
        }
        
        private ColorInfo ExtractColorFromColorType(A.ForegroundColor fgClr, SlidePart slidePart)
        {
            var rgbColor = fgClr.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null)
            {
                return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            }
            
            var schemeColor = fgClr.SchemeColor;
            if (schemeColor != null)
            {
                return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            }
            
            return new ColorInfo(0, 0, 0, true);
        }
        
        public void ApplyToShape(Shape shape, SlidePart slidePart)
        {
            if (!IsEnabled || Fill == null) return;
            
            try
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) return;
                
                // 移除现有填充
                spPr.RemoveAllChildren<A.NoFill>();
                spPr.RemoveAllChildren<A.SolidFill>();
                spPr.RemoveAllChildren<A.GradientFill>();
                spPr.RemoveAllChildren<A.PatternFill>();
                
                if (!Fill.HasFill || Fill.Color == null || Fill.Color.IsTransparent)
                {
                    spPr.AppendChild(new A.NoFill());
                }
                else
                {
                    // 使用 ColorHelper.CreateSolidFill 来保持主题色信息（无损）
                    var solidFill = ColorHelper.CreateSolidFill(Fill.Color);
                    if (solidFill != null)
                    {
                        spPr.AppendChild(solidFill);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用填充信息时出错: {ex.Message}");
            }
        }
        
        public string ToJson()
        {
            if (!IsEnabled) return "null";
            
            var jsonParts = new List<string>();
            
            // 基本颜色信息（用于预览）
            string colorStr = Fill.Color?.ToString() ?? "";
            jsonParts.Add($"\"color\":\"{colorStr}\"");
            
            float opacity = Fill.HasFill ? 1.0f : 0.0f;
            jsonParts.Add($"\"opacity\":{opacity:F1}");
            
            // 原始主题色信息（用于无损写回）
            if (Fill.Color != null && Fill.Color.IsThemeColor && !string.IsNullOrEmpty(Fill.Color.SchemeColorName))
            {
                jsonParts.Add($"\"schemeColor\":\"{Fill.Color.SchemeColorName}\"");
                
                if (Fill.Color.Transforms != null && Fill.Color.Transforms.HasTransforms)
                {
                    var transformParts = new List<string>();
                    if (Fill.Color.Transforms.LumMod.HasValue)
                        transformParts.Add($"\"lumMod\":{Fill.Color.Transforms.LumMod.Value}");
                    if (Fill.Color.Transforms.LumOff.HasValue)
                        transformParts.Add($"\"lumOff\":{Fill.Color.Transforms.LumOff.Value}");
                    if (Fill.Color.Transforms.Tint.HasValue)
                        transformParts.Add($"\"tint\":{Fill.Color.Transforms.Tint.Value}");
                    if (Fill.Color.Transforms.Shade.HasValue)
                        transformParts.Add($"\"shade\":{Fill.Color.Transforms.Shade.Value}");
                    if (Fill.Color.Transforms.SatMod.HasValue)
                        transformParts.Add($"\"satMod\":{Fill.Color.Transforms.SatMod.Value}");
                    if (Fill.Color.Transforms.SatOff.HasValue)
                        transformParts.Add($"\"satOff\":{Fill.Color.Transforms.SatOff.Value}");
                    if (Fill.Color.Transforms.Alpha.HasValue)
                        transformParts.Add($"\"alpha\":{Fill.Color.Transforms.Alpha.Value}");
                    
                    if (transformParts.Count > 0)
                        jsonParts.Add($"\"colorTransforms\":{{{string.Join(",", transformParts)}}}");
                }
            }
            else if (Fill.Color != null && !string.IsNullOrEmpty(Fill.Color.OriginalHex))
            {
                // 保存原始十六进制值
                jsonParts.Add($"\"originalHex\":\"{Fill.Color.OriginalHex}\"");
            }
            
            return string.Join(",", jsonParts);
        }

        public void FromJson(object jsonData)
        {
            try
            {
                if (jsonData == null)
                {
                    Fill = new FillInfo();
                    return;
                }

                dynamic data = jsonData;
                Fill = new FillInfo();

                try
                {
                    if (data.hasFill != null)
                        Fill.HasFill = Convert.ToInt32(data.hasFill) != 0;
                    else if (data["hasFill"] != null)
                        Fill.HasFill = Convert.ToInt32(data["hasFill"]) != 0;
                }
                catch { Fill.HasFill = false; }

                try
                {
                    string colorStr = data.color?.ToString() ?? data["color"]?.ToString();
                    if (!string.IsNullOrEmpty(colorStr))
                    {
                        Fill.Color = ColorInfo.Parse(colorStr);
                        
                        // 恢复原始主题色信息（用于无损写回）
                        string schemeColor = data.schemeColor?.ToString() ?? data["schemeColor"]?.ToString();
                        if (!string.IsNullOrEmpty(schemeColor))
                        {
                            Fill.Color.IsThemeColor = true;
                            Fill.Color.SchemeColorName = schemeColor;
                            
                            // 恢复颜色修改器
                            var transforms = data.colorTransforms ?? data["colorTransforms"];
                            if (transforms != null)
                            {
                                Fill.Color.Transforms = new ColorTransforms();
                                try { if (transforms.lumMod != null) Fill.Color.Transforms.LumMod = Convert.ToInt32(transforms.lumMod); } catch { }
                                try { if (transforms.lumOff != null) Fill.Color.Transforms.LumOff = Convert.ToInt32(transforms.lumOff); } catch { }
                                try { if (transforms.tint != null) Fill.Color.Transforms.Tint = Convert.ToInt32(transforms.tint); } catch { }
                                try { if (transforms.shade != null) Fill.Color.Transforms.Shade = Convert.ToInt32(transforms.shade); } catch { }
                                try { if (transforms.satMod != null) Fill.Color.Transforms.SatMod = Convert.ToInt32(transforms.satMod); } catch { }
                                try { if (transforms.satOff != null) Fill.Color.Transforms.SatOff = Convert.ToInt32(transforms.satOff); } catch { }
                                try { if (transforms.alpha != null) Fill.Color.Transforms.Alpha = Convert.ToInt32(transforms.alpha); } catch { }
                            }
                        }
                        else
                        {
                            // 恢复原始十六进制值
                            string originalHex = data.originalHex?.ToString() ?? data["originalHex"]?.ToString();
                            if (!string.IsNullOrEmpty(originalHex))
                            {
                                Fill.Color.OriginalHex = originalHex;
                            }
                        }
                    }
                    else
                    {
                        Fill.Color = new ColorInfo(0, 0, 0, true);
                    }
                }
                catch
                {
                    Fill.Color = new ColorInfo(0, 0, 0, true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从JSON加载填充信息时出错: {ex.Message}");
                Fill = new FillInfo();
            }
        }
    }
}
