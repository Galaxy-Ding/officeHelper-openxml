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
    public class LineComponent : IElementComponent
    {
        public string ComponentType => "Line";
        public bool IsEnabled { get; set; } = true;
        public LineStyle Line { get; set; }
        
        public LineComponent() { Line = new LineStyle(); }
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            try
            {
                Line = new LineStyle();
                var spPr = shape.ShapeProperties;
                if (spPr == null) { Line.HasOutline = false; return; }
                
                var outline = spPr.GetFirstChild<A.Outline>();
                if (outline == null) { Line.HasOutline = false; return; }
                
                var noFill = outline.GetFirstChild<A.NoFill>();
                if (noFill != null) { Line.HasOutline = false; return; }
                
                Line.HasOutline = true;
                if (outline.Width != null) Line.Weight = (float)UnitConverter.EmuToPoints(outline.Width.Value);
                
                var solidFill = outline.GetFirstChild<A.SolidFill>();
                if (solidFill != null) Line.Color = ExtractColorFromSolidFill(solidFill, slidePart);
                
                var prstDash = outline.GetFirstChild<A.PresetDash>();
                if (prstDash != null && prstDash.Val != null)
                {
                    Line.DashStyle = ConvertDashStyle(prstDash.Val.Value);
                    Line.DashStyleName = GetDashStyleName(Line.DashStyle);
                }
                else { Line.DashStyle = LineDashStyle.Solid; Line.DashStyleName = "实线"; }
                
                ExtractArrowInfo(outline);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取线条信息时出错: {ex.Message}");
                Line = new LineStyle();
            }
        }
        
        private ColorInfo ExtractColorFromSolidFill(A.SolidFill solidFill, SlidePart slidePart)
        {
            var rgbColor = solidFill.RgbColorModelHex;
            if (rgbColor != null && rgbColor.Val != null) return ColorHelper.ParseHexColor(rgbColor.Val.Value);
            var schemeColor = solidFill.SchemeColor;
            if (schemeColor != null) return ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
            return new ColorInfo(0, 0, 0, false);
        }
        
        private void ExtractArrowInfo(A.Outline outline)
        {
            var headEnd = outline.GetFirstChild<A.HeadEnd>();
            if (headEnd != null && headEnd.Type != null)
            {
                Line.HasBeginArrow = headEnd.Type.Value != A.LineEndValues.None;
                Line.BeginArrowStyle = ConvertArrowStyle(headEnd.Type.Value);
                if (headEnd.Length != null) Line.BeginArrowLength = ConvertArrowLength(headEnd.Length.Value);
                if (headEnd.Width != null) Line.BeginArrowWidth = ConvertArrowWidth(headEnd.Width.Value);
            }
            var tailEnd = outline.GetFirstChild<A.TailEnd>();
            if (tailEnd != null && tailEnd.Type != null)
            {
                Line.HasEndArrow = tailEnd.Type.Value != A.LineEndValues.None;
                Line.EndArrowStyle = ConvertArrowStyle(tailEnd.Type.Value);
                if (tailEnd.Length != null) Line.EndArrowLength = ConvertArrowLength(tailEnd.Length.Value);
                if (tailEnd.Width != null) Line.EndArrowWidth = ConvertArrowWidth(tailEnd.Width.Value);
            }
        }
        
        private LineDashStyle ConvertDashStyle(A.PresetLineDashValues val)
        {
            if (val == A.PresetLineDashValues.Solid) return LineDashStyle.Solid;
            if (val == A.PresetLineDashValues.Dot) return LineDashStyle.RoundDot;
            if (val == A.PresetLineDashValues.SystemDot) return LineDashStyle.SquareDot;
            if (val == A.PresetLineDashValues.Dash) return LineDashStyle.Dash;
            if (val == A.PresetLineDashValues.DashDot) return LineDashStyle.DashDot;
            if (val == A.PresetLineDashValues.SystemDash) return LineDashStyle.LongDash;
            if (val == A.PresetLineDashValues.SystemDashDot) return LineDashStyle.LongDashDot;
            if (val == A.PresetLineDashValues.SystemDashDotDot) return LineDashStyle.LongDashDotDot;
            return LineDashStyle.Solid;
        }
        
        private ArrowheadStyle ConvertArrowStyle(A.LineEndValues val)
        {
            if (val == A.LineEndValues.None) return ArrowheadStyle.None;
            if (val == A.LineEndValues.Triangle) return ArrowheadStyle.Triangle;
            if (val == A.LineEndValues.Stealth) return ArrowheadStyle.Stealth;
            if (val == A.LineEndValues.Diamond) return ArrowheadStyle.Diamond;
            if (val == A.LineEndValues.Oval) return ArrowheadStyle.Oval;
            if (val == A.LineEndValues.Arrow) return ArrowheadStyle.Open;
            return ArrowheadStyle.None;
        }
        
        private ArrowheadLength ConvertArrowLength(A.LineEndLengthValues val)
        {
            if (val == A.LineEndLengthValues.Small) return ArrowheadLength.Short;
            if (val == A.LineEndLengthValues.Medium) return ArrowheadLength.Medium;
            if (val == A.LineEndLengthValues.Large) return ArrowheadLength.Long;
            return ArrowheadLength.Medium;
        }
        
        private ArrowheadWidth ConvertArrowWidth(A.LineEndWidthValues val)
        {
            if (val == A.LineEndWidthValues.Small) return ArrowheadWidth.Narrow;
            if (val == A.LineEndWidthValues.Medium) return ArrowheadWidth.Medium;
            if (val == A.LineEndWidthValues.Large) return ArrowheadWidth.Wide;
            return ArrowheadWidth.Medium;
        }
        
        private string GetDashStyleName(LineDashStyle style)
        {
            if (style == LineDashStyle.Solid) return "实线";
            if (style == LineDashStyle.SquareDot) return "方形点线";
            if (style == LineDashStyle.RoundDot) return "圆形点线";
            if (style == LineDashStyle.Dash) return "虚线";
            if (style == LineDashStyle.DashDot) return "点划线";
            if (style == LineDashStyle.LongDash) return "长虚线";
            if (style == LineDashStyle.LongDashDot) return "长点划线";
            if (style == LineDashStyle.LongDashDotDot) return "长双点划线";
            return "实线";
        }
        
        private string GetArrowStyleName(ArrowheadStyle style)
        {
            if (style == ArrowheadStyle.None) return "无箭头";
            if (style == ArrowheadStyle.Triangle) return "三角形";
            if (style == ArrowheadStyle.Stealth) return "隐身箭头";
            if (style == ArrowheadStyle.Diamond) return "菱形";
            if (style == ArrowheadStyle.Oval) return "椭圆形";
            if (style == ArrowheadStyle.Open) return "开放箭头";
            return "无箭头";
        }
        
        public void ApplyToShape(Shape shape, SlidePart slidePart)
        {
            if (!IsEnabled || Line == null) return;
            try
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) return;
                spPr.RemoveAllChildren<A.Outline>();
                if (!Line.HasOutline || Line.Weight <= 0)
                {
                    var outline = new A.Outline();
                    outline.AppendChild(new A.NoFill());
                    spPr.AppendChild(outline);
                    return;
                }
                var newOutline = new A.Outline { Width = (int)UnitConverter.PointsToEmu(Line.Weight) };
                if (Line.Color != null && !Line.Color.IsTransparent)
                {
                    // 使用 ColorHelper.CreateSolidFill 来保持主题色信息（无损）
                    var solidFill = ColorHelper.CreateSolidFill(Line.Color);
                    if (solidFill != null)
                    {
                        newOutline.AppendChild(solidFill);
                    }
                }
                spPr.AppendChild(newOutline);
            }
            catch (Exception ex) { Console.WriteLine($"应用线条信息时出错: {ex.Message}"); }
        }
        
        public string ToJson()
        {
            if (!IsEnabled) return "null";
            var jsonParts = new List<string>();
            jsonParts.Add($"\"has_outline\":{(Line.HasOutline ? 1 : 0)}");
            
            // 基本颜色信息（用于预览）
            jsonParts.Add($"\"color\":\"{Line.Color?.ToString() ?? ""}\"");
            
            // 原始主题色信息（用于无损写回）
            if (Line.Color != null && Line.Color.IsThemeColor && !string.IsNullOrEmpty(Line.Color.SchemeColorName))
            {
                jsonParts.Add($"\"schemeColor\":\"{Line.Color.SchemeColorName}\"");
                
                if (Line.Color.Transforms != null && Line.Color.Transforms.HasTransforms)
                {
                    var transformParts = new List<string>();
                    if (Line.Color.Transforms.LumMod.HasValue)
                        transformParts.Add($"\"lumMod\":{Line.Color.Transforms.LumMod.Value}");
                    if (Line.Color.Transforms.LumOff.HasValue)
                        transformParts.Add($"\"lumOff\":{Line.Color.Transforms.LumOff.Value}");
                    if (Line.Color.Transforms.Tint.HasValue)
                        transformParts.Add($"\"tint\":{Line.Color.Transforms.Tint.Value}");
                    if (Line.Color.Transforms.Shade.HasValue)
                        transformParts.Add($"\"shade\":{Line.Color.Transforms.Shade.Value}");
                    if (Line.Color.Transforms.SatMod.HasValue)
                        transformParts.Add($"\"satMod\":{Line.Color.Transforms.SatMod.Value}");
                    if (Line.Color.Transforms.SatOff.HasValue)
                        transformParts.Add($"\"satOff\":{Line.Color.Transforms.SatOff.Value}");
                    if (Line.Color.Transforms.Alpha.HasValue)
                        transformParts.Add($"\"alpha\":{Line.Color.Transforms.Alpha.Value}");
                    
                    if (transformParts.Count > 0)
                        jsonParts.Add($"\"colorTransforms\":{{{string.Join(",", transformParts)}}}");
                }
            }
            else if (Line.Color != null && !string.IsNullOrEmpty(Line.Color.OriginalHex))
            {
                // 保存原始十六进制值
                jsonParts.Add($"\"originalHex\":\"{Line.Color.OriginalHex}\"");
            }
            
            jsonParts.Add($"\"width\":{Line.Weight:F2}");
            if (Line.Transparency > 0) jsonParts.Add($"\"transparency\":{Line.Transparency}");
            if (Line.DashStyle != LineDashStyle.Solid)
            {
                jsonParts.Add($"\"dash_style\":\"{Line.DashStyleName}\"");
            }
            if (Line.HasBeginArrow)
            {
                jsonParts.Add($"\"has_begin_arrow\":1");
                jsonParts.Add($"\"begin_arrow_style\":\"{GetArrowStyleName(Line.BeginArrowStyle)}\"");
            }
            if (Line.HasEndArrow)
            {
                jsonParts.Add($"\"has_end_arrow\":1");
                jsonParts.Add($"\"end_arrow_style\":\"{GetArrowStyleName(Line.EndArrowStyle)}\"");
            }
            return string.Join(",", jsonParts);
        }

        public void FromJson(object jsonData)
        {
            try
            {
                if (jsonData == null) { Line = new LineStyle(); return; }
                dynamic data = jsonData;
                Line = new LineStyle();
                try { if (data.has_outline != null) Line.HasOutline = Convert.ToInt32(data.has_outline) != 0; } catch { }
                
                try 
                { 
                    string colorStr = data.color?.ToString(); 
                    if (!string.IsNullOrEmpty(colorStr)) 
                    {
                        Line.Color = ColorInfo.Parse(colorStr);
                        
                        // 恢复原始主题色信息（用于无损写回）
                        string schemeColor = data.schemeColor?.ToString() ?? data["schemeColor"]?.ToString();
                        if (!string.IsNullOrEmpty(schemeColor))
                        {
                            Line.Color.IsThemeColor = true;
                            Line.Color.SchemeColorName = schemeColor;
                            
                            // 恢复颜色修改器
                            var transforms = data.colorTransforms ?? data["colorTransforms"];
                            if (transforms != null)
                            {
                                Line.Color.Transforms = new ColorTransforms();
                                try { if (transforms.lumMod != null) Line.Color.Transforms.LumMod = Convert.ToInt32(transforms.lumMod); } catch { }
                                try { if (transforms.lumOff != null) Line.Color.Transforms.LumOff = Convert.ToInt32(transforms.lumOff); } catch { }
                                try { if (transforms.tint != null) Line.Color.Transforms.Tint = Convert.ToInt32(transforms.tint); } catch { }
                                try { if (transforms.shade != null) Line.Color.Transforms.Shade = Convert.ToInt32(transforms.shade); } catch { }
                                try { if (transforms.satMod != null) Line.Color.Transforms.SatMod = Convert.ToInt32(transforms.satMod); } catch { }
                                try { if (transforms.satOff != null) Line.Color.Transforms.SatOff = Convert.ToInt32(transforms.satOff); } catch { }
                                try { if (transforms.alpha != null) Line.Color.Transforms.Alpha = Convert.ToInt32(transforms.alpha); } catch { }
                            }
                        }
                        else
                        {
                            // 恢复原始十六进制值
                            string originalHex = data.originalHex?.ToString() ?? data["originalHex"]?.ToString();
                            if (!string.IsNullOrEmpty(originalHex))
                            {
                                Line.Color.OriginalHex = originalHex;
                            }
                        }
                    }
                } 
                catch { }
                
                try { if (data.width != null) Line.Weight = Convert.ToSingle(data.width); } catch { }
            }
            catch (Exception ex) { Console.WriteLine($"从JSON加载线条信息时出错: {ex.Message}"); Line = new LineStyle(); }
        }
    }
}
