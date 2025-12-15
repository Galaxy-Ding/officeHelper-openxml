using System;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    /// <summary>
    /// 阴影组件 - 提取形状级别的阴影效果
    /// </summary>
    public class ShadowComponent : IElementComponent
    {
        public string ComponentType => "Shadow";
        public bool IsEnabled { get; set; } = true;
        
        public ShadowInfo Shadow { get; set; }
        
        public ShadowComponent()
        {
            Shadow = new ShadowInfo();
        }
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            try
            {
                Shadow = new ShadowInfo();
                
                var spPr = shape.ShapeProperties;
                if (spPr == null) return;
                
                // 检查 EffectList 中的阴影效果
                var effectList = spPr.GetFirstChild<A.EffectList>();
                if (effectList != null)
                {
                    // 外部阴影
                    var outerShadow = effectList.GetFirstChild<A.OuterShadow>();
                    if (outerShadow != null)
                    {
                        ExtractOuterShadow(outerShadow, slidePart);
                        return;
                    }
                    
                    // 内部阴影
                    var innerShadow = effectList.GetFirstChild<A.InnerShadow>();
                    if (innerShadow != null)
                    {
                        ExtractInnerShadow(innerShadow, slidePart);
                        return;
                    }
                }
                
                // 检查 EffectDag (Effect Diagram) 中的阴影
                var effectDag = spPr.GetFirstChild<A.EffectDag>();
                if (effectDag != null)
                {
                    var outerShadow = effectDag.GetFirstChild<A.OuterShadow>();
                    if (outerShadow != null)
                    {
                        ExtractOuterShadow(outerShadow, slidePart);
                        return;
                    }
                    
                    var innerShadow = effectDag.GetFirstChild<A.InnerShadow>();
                    if (innerShadow != null)
                    {
                        ExtractInnerShadow(innerShadow, slidePart);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取阴影信息时出错: {ex.Message}");
                Shadow = new ShadowInfo();
            }
        }
        
        private void ExtractOuterShadow(A.OuterShadow outerShadow, SlidePart slidePart)
        {
            Shadow.HasShadow = true;
            Shadow.Type = ShadowType.Outer;
            Shadow.ShadowTypeName = "outer";
            
            // 提取阴影参数
            // BlurRadius: EMU 单位，1 point = 12700 EMU
            Shadow.Blur = outerShadow.BlurRadius != null ? (float)(outerShadow.BlurRadius.Value / 12700.0) : 0;
            
            // Distance: EMU 单位
            Shadow.Distance = outerShadow.Distance != null ? (float)(outerShadow.Distance.Value / 12700.0) : 0;
            
            // Direction: 角度，单位是 1/60000 度
            Shadow.Angle = outerShadow.Direction != null ? (float)(outerShadow.Direction.Value / 60000.0) : 0;
            
            // 提取阴影颜色
            ExtractShadowColor(outerShadow, slidePart);
        }
        
        private void ExtractInnerShadow(A.InnerShadow innerShadow, SlidePart slidePart)
        {
            Shadow.HasShadow = true;
            Shadow.Type = ShadowType.Inner;
            Shadow.ShadowTypeName = "inner";
            
            // 提取阴影参数
            Shadow.Blur = innerShadow.BlurRadius != null ? (float)(innerShadow.BlurRadius.Value / 12700.0) : 0;
            Shadow.Distance = innerShadow.Distance != null ? (float)(innerShadow.Distance.Value / 12700.0) : 0;
            Shadow.Angle = innerShadow.Direction != null ? (float)(innerShadow.Direction.Value / 60000.0) : 0;
            
            // 提取阴影颜色
            ExtractShadowColor(innerShadow, slidePart);
        }
        
        private void ExtractShadowColor(DocumentFormat.OpenXml.OpenXmlElement shadowEffect, SlidePart slidePart)
        {
            try
            {
                // 尝试提取 RgbColorModelHex
                var rgbColor = shadowEffect.GetFirstChild<A.RgbColorModelHex>();
                if (rgbColor != null && rgbColor.Val != null)
                {
                    Shadow.Color = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                    
                    // 提取透明度 (Alpha)
                    var alpha = rgbColor.GetFirstChild<A.Alpha>();
                    if (alpha != null && alpha.Val != null)
                    {
                        // Alpha 值: 100000 = 100% 不透明, 0 = 完全透明
                        // Transparency: 0 = 不透明, 100 = 完全透明
                        Shadow.Transparency = 100 - (alpha.Val.Value / 1000f);
                        Shadow.Opacity = alpha.Val.Value / 1000f;
                    }
                    return;
                }
                
                // 尝试提取 SchemeColor
                var schemeColor = shadowEffect.GetFirstChild<A.SchemeColor>();
                if (schemeColor != null)
                {
                    Shadow.Color = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                    
                    // 提取透明度
                    var alpha = schemeColor.GetFirstChild<A.Alpha>();
                    if (alpha != null && alpha.Val != null)
                    {
                        Shadow.Transparency = 100 - (alpha.Val.Value / 1000f);
                        Shadow.Opacity = alpha.Val.Value / 1000f;
                    }
                    return;
                }
                
                // 默认黑色阴影
                Shadow.Color = new ColorInfo(0, 0, 0, false);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取阴影颜色时出错: {ex.Message}");
                Shadow.Color = new ColorInfo(0, 0, 0, false);
            }
        }
        
        public void ApplyToShape(Shape shape, SlidePart slidePart)
        {
            if (!IsEnabled || Shadow == null || !Shadow.HasShadow) return;
            
            try
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) return;
                
                // 移除现有的阴影效果
                spPr.RemoveAllChildren<A.EffectList>();
                spPr.RemoveAllChildren<A.EffectDag>();
                
                // 创建新的 EffectList
                var effectList = new A.EffectList();
                
                if (Shadow.Type == ShadowType.Outer)
                {
                    var outerShadow = new A.OuterShadow
                    {
                        BlurRadius = (long)(Shadow.Blur * 12700),
                        Distance = (long)(Shadow.Distance * 12700),
                        Direction = (int)(Shadow.Angle * 60000),
                        Alignment = A.RectangleAlignmentValues.BottomRight
                    };
                    
                    // 添加颜色
                    if (Shadow.Color != null && !Shadow.Color.IsTransparent)
                    {
                        var rgbColor = new A.RgbColorModelHex { Val = Shadow.Color.ToHex().TrimStart('#') };
                        
                        // 添加透明度
                        if (Shadow.Transparency > 0)
                        {
                            var alpha = new A.Alpha { Val = (int)((100 - Shadow.Transparency) * 1000) };
                            rgbColor.Append(alpha);
                        }
                        
                        outerShadow.Append(rgbColor);
                    }
                    
                    effectList.Append(outerShadow);
                }
                else if (Shadow.Type == ShadowType.Inner)
                {
                    var innerShadow = new A.InnerShadow
                    {
                        BlurRadius = (long)(Shadow.Blur * 12700),
                        Distance = (long)(Shadow.Distance * 12700),
                        Direction = (int)(Shadow.Angle * 60000)
                    };
                    
                    // 添加颜色
                    if (Shadow.Color != null && !Shadow.Color.IsTransparent)
                    {
                        var rgbColor = new A.RgbColorModelHex { Val = Shadow.Color.ToHex().TrimStart('#') };
                        
                        if (Shadow.Transparency > 0)
                        {
                            var alpha = new A.Alpha { Val = (int)((100 - Shadow.Transparency) * 1000) };
                            rgbColor.Append(alpha);
                        }
                        
                        innerShadow.Append(rgbColor);
                    }
                    
                    effectList.Append(innerShadow);
                }
                
                spPr.Append(effectList);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用阴影信息时出错: {ex.Message}");
            }
        }
        
        public string ToJson()
        {
            if (!IsEnabled || Shadow == null) return "null";
            
            return $"\"has_shadow\":{(Shadow.HasShadow ? 1 : 0)}," +
                   $"\"color\":\"{Shadow.Color?.ToString() ?? "RGB(0, 0, 0)"}\"," +
                   $"\"opacity\":{Shadow.Opacity:F1}," +
                   $"\"blur\":{Shadow.Blur:F1}," +
                   $"\"distance\":{Shadow.Distance:F1}," +
                   $"\"angle\":{Shadow.Angle:F1}," +
                   $"\"offset_x\":{Shadow.OffsetX:F1}," +
                   $"\"offset_y\":{Shadow.OffsetY:F1}," +
                   $"\"size\":{Shadow.Size:F1}," +
                   $"\"transparency\":{Shadow.Transparency:F1}," +
                   $"\"type\":{(int)Shadow.Type}," +
                   $"\"style\":{(int)Shadow.Style}," +
                   $"\"shadow_type\":\"{Shadow.ShadowTypeName}\"";
        }
        
        public void FromJson(object jsonData)
        {
            try
            {
                if (jsonData == null)
                {
                    Shadow = new ShadowInfo();
                    return;
                }
                
                dynamic data = jsonData;
                Shadow = new ShadowInfo();
                
                try { if (data.has_shadow != null) Shadow.HasShadow = Convert.ToInt32(data.has_shadow) != 0; } catch { }
                try { if (data.blur != null) Shadow.Blur = Convert.ToSingle(data.blur); } catch { }
                try { if (data.transparency != null) Shadow.Transparency = Convert.ToSingle(data.transparency); } catch { }
                try { if (data.opacity != null) Shadow.Opacity = Convert.ToSingle(data.opacity); } catch { }
                
                try
                {
                    string colorStr = data.color?.ToString();
                    if (!string.IsNullOrEmpty(colorStr))
                        Shadow.Color = ColorInfo.Parse(colorStr);
                }
                catch { }
                
                try
                {
                    string shadowType = data.shadow_type?.ToString();
                    if (shadowType == "outer") Shadow.Type = ShadowType.Outer;
                    else if (shadowType == "inner") Shadow.Type = ShadowType.Inner;
                }
                catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从JSON加载阴影信息时出错: {ex.Message}");
                Shadow = new ShadowInfo();
            }
        }
    }
}
