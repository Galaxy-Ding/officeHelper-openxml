using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    public class TextComponent : IElementComponent
    {
        public string ComponentType => "Text";
        public bool IsEnabled { get; set; } = true;
        
        public bool HasText { get; set; }
        public string TextContent { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public ColorInfo FontColor { get; set; }
        public List<ParagraphInfo> Paragraphs { get; set; }
        
        public TextComponent()
        {
            HasText = false;
            TextContent = "";
            FontName = "";
            FontSize = 12;
            FontColor = new ColorInfo();
            Paragraphs = new List<ParagraphInfo>();
        }
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            try
            {
                Paragraphs = new List<ParagraphInfo>();
                var textBody = shape.TextBody;
                if (textBody == null) { HasText = false; return; }
                
                var sb = new StringBuilder();
                bool firstPara = true;
                
                foreach (var para in textBody.Elements<A.Paragraph>())
                {
                    var paraInfo = ExtractParagraph(para, slidePart);
                    Paragraphs.Add(paraInfo);
                    
                    if (!firstPara) sb.Append("\n");
                    sb.Append(paraInfo.GetPlainText());
                    firstPara = false;
                }
                
                TextContent = sb.ToString();
                HasText = !string.IsNullOrEmpty(TextContent);
                
                // 提取第一个运行的字体信息作为默认
                if (Paragraphs.Count > 0 && Paragraphs[0].Runs.Count > 0)
                {
                    var firstRun = Paragraphs[0].Runs[0];
                    FontName = firstRun.FontName;
                    FontSize = firstRun.FontSize;
                    FontColor = firstRun.FontColor;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取文本信息时出错: {ex.Message}");
                HasText = false;
                TextContent = "";
            }
        }
        
        private ParagraphInfo ExtractParagraph(A.Paragraph para, SlidePart slidePart)
        {
            var paraInfo = new ParagraphInfo();

            // 提取段落属性
            var pPr = para.ParagraphProperties;
            if (pPr != null)
            {
                if (pPr.Alignment != null)
                    paraInfo.Alignment = ConvertAlignment(pPr.Alignment.Value);
                if (pPr.Level != null)
                    paraInfo.Level = pPr.Level.Value;
            }

            // 按顺序提取文本运行和换行符
            foreach (var child in para.ChildElements)
            {
                if (child is A.Run run)
                {
                    var runInfo = ExtractRun(run, slidePart);
                    if (!string.IsNullOrEmpty(runInfo.Text))
                        paraInfo.Runs.Add(runInfo);
                }
                else if (child is A.Break)
                {
                    // 段内换行符 - 添加一个包含\n的Run
                    var brInfo = new TextRunInfo { Text = "\n" };
                    paraInfo.Runs.Add(brInfo);
                }
            }

            return paraInfo;
        }
        
        private TextRunInfo ExtractRun(A.Run run, SlidePart slidePart)
        {
            var runInfo = new TextRunInfo();

            // 提取文本
            var text = run.GetFirstChild<A.Text>();
            runInfo.Text = text?.Text ?? "";

            // 提取运行属性
            var rPr = run.RunProperties;
            if (rPr != null)
            {
                // 字体大小 (百分之一磅)
                if (rPr.FontSize != null)
                    runInfo.FontSize = rPr.FontSize.Value / 100f;

                // 粗体
                if (rPr.Bold != null)
                    runInfo.IsBold = rPr.Bold.Value;

                // 斜体
                if (rPr.Italic != null)
                    runInfo.IsItalic = rPr.Italic.Value;

                // 下划线
                if (rPr.Underline != null)
                    runInfo.IsUnderline = rPr.Underline.Value != A.TextUnderlineValues.None;

                // 删除线
                if (rPr.Strike != null)
                    runInfo.IsStrikethrough = rPr.Strike.Value != A.TextStrikeValues.NoStrike;

                // 字体颜色
                var solidFill = rPr.GetFirstChild<A.SolidFill>();
                if (solidFill != null)
                    runInfo.FontColor = ExtractColorFromSolidFill(solidFill, slidePart);

                // 字体名称 - 优先使用东亚字体（中文），然后拉丁字体
                string fontName = null;

                var ea = rPr.GetFirstChild<A.EastAsianFont>();
                if (ea != null && ea.Typeface != null)
                    fontName = ea.Typeface;

                var latin = rPr.GetFirstChild<A.LatinFont>();
                if (latin != null && latin.Typeface != null && string.IsNullOrEmpty(fontName))
                    fontName = latin.Typeface;

                // 解析主题字体引用
                runInfo.FontName = ResolveThemeFont(fontName, slidePart);

                // 提取文字填充效果 (WordArt)
                runInfo.TextFill = ExtractTextFill(rPr, slidePart);
                
                // 提取文字轮廓效果 (WordArt)
                runInfo.TextOutline = ExtractTextOutline(rPr, slidePart);
                
                // 提取文字效果 (WordArt) - includes shadow, glow, reflection, soft edges
                runInfo.TextEffects = ExtractTextEffects(rPr, slidePart);
                
                // Update legacy shadow properties for backward compatibility
                if (runInfo.TextEffects.HasShadow && runInfo.TextEffects.Shadow != null)
                {
                    runInfo.HasShadow = true;
                    runInfo.Shadow = runInfo.TextEffects.Shadow;
                }
            }

            return runInfo;
        }



        /// <summary>
        /// 提取文字填充效果 (WordArt text fill)
        /// Extracts solid fill, gradient fill, pattern fill, and no fill from RunProperties
        /// </summary>
        private TextFillInfo ExtractTextFill(A.RunProperties rPr, SlidePart slidePart)
        {
            var textFill = new TextFillInfo();

            if (rPr == null)
            {
                textFill.HasFill = false;
                textFill.FillType = FillType.NoFill;
                return textFill;
            }

            try
            {
                // Check for NoFill
                var noFill = rPr.GetFirstChild<A.NoFill>();
                if (noFill != null)
                {
                    textFill.HasFill = false;
                    textFill.FillType = FillType.NoFill;
                    return textFill;
                }

                // Check for SolidFill
                var solidFill = rPr.GetFirstChild<A.SolidFill>();
                if (solidFill != null)
                {
                    textFill.HasFill = true;
                    textFill.FillType = FillType.Solid;
                    textFill.Color = ColorHelper.ExtractColorInfo(solidFill, slidePart);
                    
                    // Extract transparency from alpha if present
                    var rgbColor = solidFill.RgbColorModelHex;
                    if (rgbColor != null)
                    {
                        var alpha = rgbColor.GetFirstChild<A.Alpha>();
                        if (alpha?.Val != null)
                        {
                            textFill.Transparency = 1.0f - (alpha.Val.Value / 100000.0f);
                        }
                    }
                    
                    var schemeColor = solidFill.SchemeColor;
                    if (schemeColor != null)
                    {
                        var alpha = schemeColor.GetFirstChild<A.Alpha>();
                        if (alpha?.Val != null)
                        {
                            textFill.Transparency = 1.0f - (alpha.Val.Value / 100000.0f);
                        }
                    }
                    
                    return textFill;
                }

                // Check for GradientFill
                var gradFill = rPr.GetFirstChild<A.GradientFill>();
                if (gradFill != null)
                {
                    textFill.HasFill = true;
                    textFill.FillType = FillType.Gradient;
                    textFill.Gradient = ExtractGradientInfo(gradFill, slidePart);
                    return textFill;
                }

                // Check for PatternFill
                var pattFill = rPr.GetFirstChild<A.PatternFill>();
                if (pattFill != null)
                {
                    textFill.HasFill = true;
                    textFill.FillType = FillType.Pattern;
                    textFill.Pattern = ExtractPatternInfo(pattFill, slidePart);
                    return textFill;
                }

                // If no fill type is specified, default to no fill
                textFill.HasFill = false;
                textFill.FillType = FillType.NoFill;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取文字填充时出错: {ex.Message}");
                textFill.HasFill = false;
                textFill.FillType = FillType.NoFill;
            }

            return textFill;
        }

        /// <summary>
        /// Extract gradient fill information
        /// </summary>
        private GradientInfo ExtractGradientInfo(A.GradientFill gradFill, SlidePart slidePart)
        {
            var gradInfo = new GradientInfo();

            try
            {
                // Extract gradient stops
                var gsLst = gradFill.GradientStopList;
                if (gsLst != null)
                {
                    foreach (var gs in gsLst.Elements<A.GradientStop>())
                    {
                        if (gs.Position != null)
                        {
                            float position = gs.Position.Value / 100000.0f;
                            ColorInfo color = new ColorInfo();

                            // Extract color from gradient stop
                            var solidFill = gs.GetFirstChild<A.RgbColorModelHex>();
                            if (solidFill != null && solidFill.Val != null)
                            {
                                color = ColorHelper.ParseHexColor(solidFill.Val.Value);
                                color.OriginalHex = solidFill.Val.Value;
                            }

                            var schemeColor = gs.GetFirstChild<A.SchemeColor>();
                            if (schemeColor != null)
                            {
                                color = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                            }

                            gradInfo.Stops.Add(new GradientStop(position, color));
                        }
                    }
                }

                // Extract gradient type and angle
                var lin = gradFill.GetFirstChild<A.LinearGradientFill>();
                if (lin != null)
                {
                    gradInfo.GradientType = "Linear";
                    if (lin.Angle != null)
                    {
                        gradInfo.Angle = lin.Angle.Value / 60000.0f;
                    }
                }

                var path = gradFill.GetFirstChild<A.PathGradientFill>();
                if (path != null)
                {
                    gradInfo.GradientType = path.Path?.Value.ToString() ?? "Path";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取渐变填充时出错: {ex.Message}");
            }

            return gradInfo;
        }

        /// <summary>
        /// Extract pattern fill information
        /// </summary>
        private PatternInfo ExtractPatternInfo(A.PatternFill pattFill, SlidePart slidePart)
        {
            var pattInfo = new PatternInfo();

            try
            {
                // Extract pattern type
                if (pattFill.Preset != null)
                {
                    pattInfo.PatternType = pattFill.Preset.Value.ToString();
                }

                // Extract foreground color
                var fgColor = pattFill.ForegroundColor;
                if (fgColor != null)
                {
                    var rgbColor = fgColor.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        pattInfo.ForegroundColor = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                    }

                    var schemeColor = fgColor.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        pattInfo.ForegroundColor = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                    }
                }

                // Extract background color
                var bgColor = pattFill.BackgroundColor;
                if (bgColor != null)
                {
                    var rgbColor = bgColor.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        pattInfo.BackgroundColor = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                    }

                    var schemeColor = bgColor.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        pattInfo.BackgroundColor = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取图案填充时出错: {ex.Message}");
            }

            return pattInfo;
        }

        /// <summary>
        /// Extract text outline information (WordArt text outline)
        /// Extracts outline properties from RunProperties including width, color, dash style, compound type, cap type, and join type
        /// </summary>
        private TextOutlineInfo ExtractTextOutline(A.RunProperties rPr, SlidePart slidePart)
        {
            var textOutline = new TextOutlineInfo();

            if (rPr == null)
            {
                textOutline.HasOutline = false;
                return textOutline;
            }

            try
            {
                // Check for Outline element (a:ln)
                var outline = rPr.GetFirstChild<A.Outline>();
                if (outline == null)
                {
                    textOutline.HasOutline = false;
                    return textOutline;
                }

                textOutline.HasOutline = true;

                // Extract outline width (in EMUs, convert to points)
                // 1 point = 12700 EMUs
                if (outline.Width != null)
                {
                    textOutline.Width = outline.Width.Value / 12700.0f;
                }

                // Extract outline color from SolidFill
                var solidFill = outline.GetFirstChild<A.SolidFill>();
                if (solidFill != null)
                {
                    textOutline.Color = ColorHelper.ExtractColorInfo(solidFill, slidePart);
                    
                    // Extract transparency from alpha if present
                    var rgbColor = solidFill.RgbColorModelHex;
                    if (rgbColor != null)
                    {
                        var alpha = rgbColor.GetFirstChild<A.Alpha>();
                        if (alpha?.Val != null)
                        {
                            textOutline.Transparency = 1.0f - (alpha.Val.Value / 100000.0f);
                        }
                    }
                    
                    var schemeColor = solidFill.SchemeColor;
                    if (schemeColor != null)
                    {
                        var alpha = schemeColor.GetFirstChild<A.Alpha>();
                        if (alpha?.Val != null)
                        {
                            textOutline.Transparency = 1.0f - (alpha.Val.Value / 100000.0f);
                        }
                    }
                }

                // Extract dash style
                var prstDash = outline.GetFirstChild<A.PresetDash>();
                if (prstDash?.Val != null)
                {
                    textOutline.DashStyle = ConvertDashStyle(prstDash.Val.Value);
                }

                // Extract compound line type
                if (outline.CompoundLineType != null)
                {
                    textOutline.CompoundLineType = outline.CompoundLineType.Value.ToString();
                }

                // Extract cap type
                if (outline.CapType != null)
                {
                    textOutline.CapType = outline.CapType.Value.ToString();
                }

                // Extract join type
                var round = outline.GetFirstChild<A.Round>();
                var bevel = outline.GetFirstChild<A.Bevel>();
                var miter = outline.GetFirstChild<A.Miter>();
                
                if (round != null)
                {
                    textOutline.JoinType = "Round";
                }
                else if (bevel != null)
                {
                    textOutline.JoinType = "Bevel";
                }
                else if (miter != null)
                {
                    textOutline.JoinType = "Miter";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取文字轮廓时出错: {ex.Message}");
                textOutline.HasOutline = false;
            }

            return textOutline;
        }

        /// <summary>
        /// Convert OpenXML PresetLineDashValues to LineDashStyle enum
        /// </summary>
        private LineDashStyle ConvertDashStyle(A.PresetLineDashValues dashValue)
        {
            if (dashValue == A.PresetLineDashValues.Solid) return LineDashStyle.Solid;
            if (dashValue == A.PresetLineDashValues.Dot) return LineDashStyle.RoundDot;
            if (dashValue == A.PresetLineDashValues.SystemDot) return LineDashStyle.SquareDot;
            if (dashValue == A.PresetLineDashValues.Dash) return LineDashStyle.Dash;
            if (dashValue == A.PresetLineDashValues.DashDot) return LineDashStyle.DashDot;
            if (dashValue == A.PresetLineDashValues.SystemDash) return LineDashStyle.LongDash;
            if (dashValue == A.PresetLineDashValues.SystemDashDot) return LineDashStyle.LongDashDot;
            if (dashValue == A.PresetLineDashValues.SystemDashDotDot) return LineDashStyle.LongDashDotDot;
            return LineDashStyle.Solid;
        }

        /// <summary>
        /// Extract text effects (WordArt text effects)
        /// Extracts shadow, glow, reflection, and soft edges from RunProperties EffectList
        /// Handles multiple effects on the same text run
        /// </summary>
        private TextEffectsInfo ExtractTextEffects(A.RunProperties rPr, SlidePart slidePart)
        {
            var textEffects = new TextEffectsInfo();

            if (rPr == null)
            {
                textEffects.HasEffects = false;
                return textEffects;
            }

            try
            {
                // Check for EffectList element
                var effectList = rPr.GetFirstChild<A.EffectList>();
                if (effectList == null)
                {
                    textEffects.HasEffects = false;
                    return textEffects;
                }

                // Extract outer shadow
                var outerShadow = effectList.GetFirstChild<A.OuterShadow>();
                if (outerShadow != null)
                {
                    textEffects.HasShadow = true;
                    textEffects.Shadow = new ShadowInfo
                    {
                        HasShadow = true,
                        Type = ShadowType.Outer,
                        Blur = outerShadow.BlurRadius != null ? (float)(outerShadow.BlurRadius.Value / 914400.0) : 0,
                        Distance = outerShadow.Distance != null ? (float)(outerShadow.Distance.Value / 914400.0) : 0,
                        Angle = outerShadow.Direction != null ? (float)(outerShadow.Direction.Value / 60000.0) : 0,
                        Transparency = 0
                    };

                    // Extract shadow color
                    var rgbColor = outerShadow.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        textEffects.Shadow.Color = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                        // Extract transparency
                        var alpha = rgbColor.GetFirstChild<A.Alpha>();
                        if (alpha != null && alpha.Val != null)
                            textEffects.Shadow.Transparency = (100 - alpha.Val.Value / 1000f);
                    }

                    var schemeColor = outerShadow.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        textEffects.Shadow.Color = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                        var alpha = schemeColor.GetFirstChild<A.Alpha>();
                        if (alpha != null && alpha.Val != null)
                            textEffects.Shadow.Transparency = (100 - alpha.Val.Value / 1000f);
                    }
                }

                // Extract inner shadow
                var innerShadow = effectList.GetFirstChild<A.InnerShadow>();
                if (innerShadow != null)
                {
                    textEffects.HasShadow = true;
                    textEffects.Shadow = new ShadowInfo
                    {
                        HasShadow = true,
                        Type = ShadowType.Inner,
                        Blur = innerShadow.BlurRadius != null ? (float)(innerShadow.BlurRadius.Value / 914400.0) : 0,
                        Distance = innerShadow.Distance != null ? (float)(innerShadow.Distance.Value / 914400.0) : 0,
                        Angle = innerShadow.Direction != null ? (float)(innerShadow.Direction.Value / 60000.0) : 0,
                        Transparency = 0
                    };

                    // Extract shadow color
                    var rgbColor = innerShadow.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        textEffects.Shadow.Color = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                        var alpha = rgbColor.GetFirstChild<A.Alpha>();
                        if (alpha != null && alpha.Val != null)
                            textEffects.Shadow.Transparency = (100 - alpha.Val.Value / 1000f);
                    }

                    var schemeColor = innerShadow.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        textEffects.Shadow.Color = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                        var alpha = schemeColor.GetFirstChild<A.Alpha>();
                        if (alpha != null && alpha.Val != null)
                            textEffects.Shadow.Transparency = (100 - alpha.Val.Value / 1000f);
                    }
                }

                // Extract glow effect
                var glow = effectList.GetFirstChild<A.Glow>();
                if (glow != null)
                {
                    textEffects.HasGlow = true;
                    textEffects.Glow = new GlowInfo
                    {
                        Radius = glow.Radius != null ? (float)(glow.Radius.Value / 12700.0) : 0,
                        Transparency = 0
                    };

                    // Extract glow color
                    var rgbColor = glow.GetFirstChild<A.RgbColorModelHex>();
                    if (rgbColor != null && rgbColor.Val != null)
                    {
                        textEffects.Glow.Color = ColorHelper.ParseHexColor(rgbColor.Val.Value);
                        var alpha = rgbColor.GetFirstChild<A.Alpha>();
                        if (alpha?.Val != null)
                        {
                            textEffects.Glow.Transparency = 1.0f - (alpha.Val.Value / 100000.0f);
                        }
                    }

                    var schemeColor = glow.GetFirstChild<A.SchemeColor>();
                    if (schemeColor != null)
                    {
                        textEffects.Glow.Color = ColorHelper.ResolveSchemeColor(schemeColor, slidePart);
                        var alpha = schemeColor.GetFirstChild<A.Alpha>();
                        if (alpha?.Val != null)
                        {
                            textEffects.Glow.Transparency = 1.0f - (alpha.Val.Value / 100000.0f);
                        }
                    }
                }

                // Extract reflection effect
                var reflection = effectList.GetFirstChild<A.Reflection>();
                if (reflection != null)
                {
                    textEffects.HasReflection = true;
                    textEffects.Reflection = new ReflectionInfo
                    {
                        BlurRadius = reflection.BlurRadius != null ? (float)(reflection.BlurRadius.Value / 12700.0) : 0,
                        StartOpacity = reflection.StartOpacity != null ? (reflection.StartOpacity.Value / 100000.0f) : 1.0f,
                        StartPosition = reflection.StartPosition != null ? (reflection.StartPosition.Value / 100000.0f) : 0.0f,
                        EndAlpha = reflection.EndAlpha != null ? (reflection.EndAlpha.Value / 100000.0f) : 0.0f,
                        EndPosition = reflection.EndPosition != null ? (reflection.EndPosition.Value / 100000.0f) : 1.0f,
                        Distance = reflection.Distance != null ? (float)(reflection.Distance.Value / 12700.0) : 0,
                        Direction = reflection.Direction != null ? (float)(reflection.Direction.Value / 60000.0) : 0,
                        FadeDirection = reflection.FadeDirection != null ? (float)(reflection.FadeDirection.Value / 60000.0) : 0,
                        SkewHorizontal = reflection.HorizontalSkew != null ? (float)(reflection.HorizontalSkew.Value / 60000.0) : 0,
                        SkewVertical = reflection.VerticalSkew != null ? (float)(reflection.VerticalSkew.Value / 60000.0) : 0
                    };
                }

                // Extract soft edges effect
                var softEdge = effectList.GetFirstChild<A.SoftEdge>();
                if (softEdge != null)
                {
                    textEffects.HasSoftEdge = true;
                    textEffects.SoftEdgeRadius = softEdge.Radius != null ? (float)(softEdge.Radius.Value / 12700.0) : 0;
                }

                // Set HasEffects flag if any effect is present
                textEffects.HasEffects = textEffects.HasShadow || textEffects.HasGlow || 
                                         textEffects.HasReflection || textEffects.HasSoftEdge;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取文字效果时出错: {ex.Message}");
                textEffects.HasEffects = false;
            }

            return textEffects;
        }

        private ColorInfo ExtractColorFromSolidFill(A.SolidFill solidFill, SlidePart slidePart)
        {
            // 使用 ColorHelper.ExtractColorInfo 来保留原始主题色信息
            return ColorHelper.ExtractColorInfo(solidFill, slidePart);
        }

        /// <summary>
        /// 对RGB颜色应用颜色变换
        /// 根据 ECMA-376 规范：
        /// - Tint/Shade 直接在 RGB 空间操作
        /// - LumMod/LumOff 在 HSL 空间操作 Luminance
        /// </summary>
        private void ApplyColorTransformsToRgb(A.RgbColorModelHex rgbColor, ref ColorInfo color)
        {
            // 检查是否有Alpha（透明度）
            var alpha = rgbColor.GetFirstChild<A.Alpha>();
            if (alpha?.Val != null && alpha.Val.Value < 100000)
            {
                // 如果透明度小于100%，可以在这里处理
                // 目前暂不处理透明度
            }

            // 检查亮度修改
            var lumMod = rgbColor.GetFirstChild<A.LuminanceModulation>();
            var lumOff = rgbColor.GetFirstChild<A.LuminanceOffset>();
            var tint = rgbColor.GetFirstChild<A.Tint>();
            var shade = rgbColor.GetFirstChild<A.Shade>();

            if (lumMod != null || lumOff != null || tint != null || shade != null)
            {
                int r = color.Red, g = color.Green, b = color.Blue;

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
                    RgbToHsl(r, g, b, out double h, out double s, out double l);

                    if (lumMod?.Val != null)
                        l *= lumMod.Val.Value / 100000.0;
                    if (lumOff?.Val != null)
                        l += lumOff.Val.Value / 100000.0;

                    l = Math.Max(0, Math.Min(1, l));
                    HslToRgb(h, s, l, out r, out g, out b);
                }

                // 确保值在有效范围内
                color.Red = Math.Max(0, Math.Min(255, r));
                color.Green = Math.Max(0, Math.Min(255, g));
                color.Blue = Math.Max(0, Math.Min(255, b));
            }
        }

        #region HSL转换辅助方法
        private void RgbToHsl(int r, int g, int b, out double h, out double s, out double l)
        {
            double rd = r / 255.0;
            double gd = g / 255.0;
            double bd = b / 255.0;

            double max = Math.Max(rd, Math.Max(gd, bd));
            double min = Math.Min(rd, Math.Min(gd, bd));

            l = (max + min) / 2.0;

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
        private void HslToRgb(double h, double s, double l, out int r, out int g, out int b)
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

        private double HueToRgb(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 1.0 / 2) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }
        #endregion

        /// <summary>
        /// 解析主题字体引用，返回实际字体名称
        /// +mj-lt = Major Latin (标题拉丁字体)
        /// +mn-lt = Minor Latin (正文拉丁字体)
        /// +mj-ea = Major East Asian (标题东亚字体)
        /// +mn-ea = Minor East Asian (正文东亚字体)
        /// </summary>
        private string ResolveThemeFont(string fontName, SlidePart slidePart)
        {
            // 如果字体名为空，尝试获取默认主题字体
            if (string.IsNullOrEmpty(fontName))
            {
                return GetDefaultThemeFont(slidePart) ?? "宋体";
            }

            // 如果不是主题字体引用，直接返回
            if (!fontName.StartsWith("+"))
            {
                return fontName;
            }

            try
            {
                // 获取主题字体方案
                var themePart = slidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart;
                if (themePart?.Theme?.ThemeElements?.FontScheme == null)
                {
                    return GetDefaultFontForThemeRef(fontName);
                }

                var fontScheme = themePart.Theme.ThemeElements.FontScheme;
                //"+mj-ea"
                switch (fontName.ToLower())
                {
                    case "+mj-lt": // Major Latin
                        var majorLatin = fontScheme.MajorFont?.LatinFont?.Typeface;
                        return string.IsNullOrEmpty(majorLatin) ? "Calibri Light" : majorLatin;
                    case "+mn-lt": // Minor Latin
                        var minorLatin = fontScheme.MinorFont?.LatinFont?.Typeface;
                        return string.IsNullOrEmpty(minorLatin) ? "Calibri" : minorLatin;
                    case "+mj-ea": // Major East Asian
                        var majorEa = fontScheme.MajorFont?.EastAsianFont?.Typeface;
                        return string.IsNullOrEmpty(majorEa) ? "等线 Light" : majorEa;
                    case "+mn-ea": // Minor East Asian
                        var minorEa = fontScheme.MinorFont?.EastAsianFont?.Typeface;
                        return string.IsNullOrEmpty(minorEa) ? "等线" : minorEa;
                    default:
                        return GetDefaultFontForThemeRef(fontName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解析主题字体时出错: {ex.Message}");
                return GetDefaultFontForThemeRef(fontName);
            }
        }

        /// <summary>
        /// 获取主题字体引用的默认字体（当无法获取主题时）
        /// </summary>
        private string GetDefaultFontForThemeRef(string fontRef)
        {
            switch (fontRef?.ToLower())
            {
                case "+mj-lt": return "Calibri Light";
                case "+mn-lt": return "Calibri";
                case "+mj-ea": return "等线 Light";
                case "+mn-ea": return "等线";
                default: return "宋体";
            }
        }

        /// <summary>
        /// 获取默认主题字体（当Run没有指定字体时）
        /// </summary>
        private string GetDefaultThemeFont(SlidePart slidePart)
        {
            try
            {
                var themePart = slidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart;
                if (themePart?.Theme?.ThemeElements?.FontScheme != null)
                {
                    var fontScheme = themePart.Theme.ThemeElements.FontScheme;
                    // 优先返回东亚字体（中文）
                    var eaFont = fontScheme.MinorFont?.EastAsianFont?.Typeface;
                    if (!string.IsNullOrEmpty(eaFont))
                        return eaFont;
                    // 其次返回拉丁字体
                    var latinFont = fontScheme.MinorFont?.LatinFont?.Typeface;
                    if (!string.IsNullOrEmpty(latinFont))
                        return latinFont;
                }
            }
            catch { }
            return null;
        }
        
        private TextAlignment ConvertAlignment(A.TextAlignmentTypeValues val)
        {
            if (val == A.TextAlignmentTypeValues.Left) return TextAlignment.Left;
            if (val == A.TextAlignmentTypeValues.Center) return TextAlignment.Center;
            if (val == A.TextAlignmentTypeValues.Right) return TextAlignment.Right;
            if (val == A.TextAlignmentTypeValues.Justified) return TextAlignment.Justify;
            if (val == A.TextAlignmentTypeValues.Distributed) return TextAlignment.Distributed;
            return TextAlignment.Left;
        }
        
        public void ApplyToShape(Shape shape, SlidePart slidePart)
        {
            if (!IsEnabled) return;
            
            try
            {
                var textBody = shape.TextBody;
                if (textBody == null)
                {
                    textBody = new TextBody();
                    shape.AppendChild(textBody);
                }
                
                // 清除现有段落
                textBody.RemoveAllChildren<A.Paragraph>();
                
                if (Paragraphs != null && Paragraphs.Count > 0)
                {
                    foreach (var paraInfo in Paragraphs)
                    {
                        var para = CreateParagraph(paraInfo);
                        textBody.AppendChild(para);
                    }
                }
                else if (!string.IsNullOrEmpty(TextContent))
                {
                    // 简单文本
                    var para = new A.Paragraph();
                    var run = new A.Run();
                    run.AppendChild(new A.Text(TextContent));
                    para.AppendChild(run);
                    textBody.AppendChild(para);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"应用文本信息时出错: {ex.Message}");
            }
        }
        
        private A.Paragraph CreateParagraph(ParagraphInfo paraInfo)
        {
            var para = new A.Paragraph();
            
            foreach (var runInfo in paraInfo.Runs)
            {
                var run = new A.Run();
                var rPr = new A.RunProperties();
                
                if (runInfo.FontSize > 0)
                    rPr.FontSize = (int)(runInfo.FontSize * 100);
                if (runInfo.IsBold)
                    rPr.Bold = true;
                if (runInfo.IsItalic)
                    rPr.Italic = true;
                
                if (runInfo.FontColor != null && !runInfo.FontColor.IsTransparent)
                {
                    // 使用 ColorHelper.CreateSolidFill 来保持主题色信息（无损）
                    var solidFill = ColorHelper.CreateSolidFill(runInfo.FontColor);
                    if (solidFill != null)
                    {
                        rPr.AppendChild(solidFill);
                    }
                }
                
                if (!string.IsNullOrEmpty(runInfo.FontName))
                    rPr.AppendChild(new A.LatinFont { Typeface = runInfo.FontName });
                
                run.AppendChild(rPr);
                run.AppendChild(new A.Text(runInfo.Text));
                para.AppendChild(run);
            }
            
            return para;
        }
        
        public string ToJson()
        {
            if (!IsEnabled) return "null";

            var sb = new StringBuilder();
            sb.Append($"\"hastext\":{(HasText ? 1 : 0)}");

            // 输出文本运行数组 (合并相同格式的连续Run)
            sb.Append(",\"text\":");
            if (Paragraphs != null && Paragraphs.Count > 0)
            {
                // 收集所有Run并合并相同格式的连续Run
                var mergedRuns = GetMergedRuns();

                sb.Append("[");
                bool first = true;
                foreach (var run in mergedRuns)
                {
                    if (!first) sb.Append(",");
                    sb.Append("{");
                    sb.Append($"\"content\":\"{EscapeJson(run.Text)}\",");
                    sb.Append($"\"font\":\"{EscapeJson(run.FontName)}\",");
                    sb.Append($"\"font_size\":{run.FontSize:F1},");
                    sb.Append($"\"font_color\":\"{run.FontColor?.ToString() ?? "RGB(0, 0, 0)"}\",");
                    
                    // 添加原始主题色信息（用于无损写回）
                    if (run.FontColor != null && run.FontColor.IsThemeColor && !string.IsNullOrEmpty(run.FontColor.SchemeColorName))
                    {
                        sb.Append($"\"schemeColor\":\"{run.FontColor.SchemeColorName}\",");
                        
                        if (run.FontColor.Transforms != null && run.FontColor.Transforms.HasTransforms)
                        {
                            sb.Append("\"colorTransforms\":{");
                            var transformParts = new List<string>();
                            if (run.FontColor.Transforms.LumMod.HasValue)
                                transformParts.Add($"\"lumMod\":{run.FontColor.Transforms.LumMod.Value}");
                            if (run.FontColor.Transforms.LumOff.HasValue)
                                transformParts.Add($"\"lumOff\":{run.FontColor.Transforms.LumOff.Value}");
                            if (run.FontColor.Transforms.Tint.HasValue)
                                transformParts.Add($"\"tint\":{run.FontColor.Transforms.Tint.Value}");
                            if (run.FontColor.Transforms.Shade.HasValue)
                                transformParts.Add($"\"shade\":{run.FontColor.Transforms.Shade.Value}");
                            if (run.FontColor.Transforms.SatMod.HasValue)
                                transformParts.Add($"\"satMod\":{run.FontColor.Transforms.SatMod.Value}");
                            if (run.FontColor.Transforms.SatOff.HasValue)
                                transformParts.Add($"\"satOff\":{run.FontColor.Transforms.SatOff.Value}");
                            if (run.FontColor.Transforms.Alpha.HasValue)
                                transformParts.Add($"\"alpha\":{run.FontColor.Transforms.Alpha.Value}");
                            sb.Append(string.Join(",", transformParts));
                            sb.Append("},");
                        }
                    }
                    else if (run.FontColor != null && !string.IsNullOrEmpty(run.FontColor.OriginalHex))
                    {
                        sb.Append($"\"originalHex\":\"{run.FontColor.OriginalHex}\",");
                    }
                    
                    // 始终输出所有格式字段
                    sb.Append($"\"font_bold\":{(run.IsBold ? 1 : 0)},");
                    sb.Append($"\"font_italic\":{(run.IsItalic ? 1 : 0)},");
                    sb.Append($"\"font_underline\":{(run.IsUnderline ? 1 : 0)},");
                    sb.Append($"\"font_strikethrough\":{(run.IsStrikethrough ? 1 : 0)}");
                    
                    // Add text fill (WordArt) - only if it has fill
                    if (run.TextFill != null && run.TextFill.HasFill)
                    {
                        sb.Append(",\"text_fill\":");
                        sb.Append(SerializeTextFillToJson(run.TextFill));
                    }
                    
                    // Add text outline (WordArt) - only if it has outline
                    if (run.TextOutline != null && run.TextOutline.HasOutline)
                    {
                        sb.Append(",\"text_outline\":");
                        sb.Append(SerializeTextOutlineToJson(run.TextOutline));
                    }
                    
                    // Add text effects (WordArt) - only if it has actual effects content
                    if (run.TextEffects != null && run.TextEffects.HasEffects)
                    {
                        sb.Append(",\"text_effects\":");
                        sb.Append(SerializeTextEffectsToJson(run.TextEffects));
                    }
                    
                    sb.Append("}");
                    first = false;
                }
                sb.Append("]");
            }
            else
            {
                sb.Append("[]");
            }

            return sb.ToString();
        }

        /// <summary>
        /// 获取合并后的Run列表（相同格式的连续Run会被合并，段落间用\n分隔）
        /// </summary>
        private List<TextRunInfo> GetMergedRuns()
        {
            var result = new List<TextRunInfo>();
            bool isFirstParagraph = true;

            foreach (var para in Paragraphs)
            {
                // 段落之间添加换行符（第一个段落除外）
                if (!isFirstParagraph && para.Runs.Count > 0)
                {
                    // 在上一个Run的末尾添加换行符
                    if (result.Count > 0)
                    {
                        result[result.Count - 1].Text += "\n";
                    }
                }
                isFirstParagraph = false;

                foreach (var run in para.Runs)
                {
                    if (result.Count > 0 && IsSameFormat(result[result.Count - 1], run))
                    {
                        // 合并文本内容
                        result[result.Count - 1].Text += run.Text;
                    }
                    else
                    {
                        // 创建新的Run副本
                        var newRun = new TextRunInfo
                        {
                            Text = run.Text,
                            FontName = run.FontName,
                            FontSize = run.FontSize,
                            IsBold = run.IsBold,
                            IsItalic = run.IsItalic,
                            IsUnderline = run.IsUnderline,
                            IsStrikethrough = run.IsStrikethrough,
                            FontColor = run.FontColor,
                            HasShadow = run.HasShadow,
                            Shadow = run.Shadow,
                            TextFill = run.TextFill,
                            TextOutline = run.TextOutline,
                            TextEffects = run.TextEffects
                        };
                        result.Add(newRun);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 判断两个Run的格式是否相同
        /// </summary>
        private bool IsSameFormat(TextRunInfo run1, TextRunInfo run2)
        {
            if (run1 == null || run2 == null) return false;

            // 比较主要格式属性
            return run1.FontName == run2.FontName &&
                   Math.Abs(run1.FontSize - run2.FontSize) < 0.1f &&
                   run1.IsBold == run2.IsBold &&
                   run1.IsItalic == run2.IsItalic &&
                   run1.IsUnderline == run2.IsUnderline &&
                   run1.IsStrikethrough == run2.IsStrikethrough &&
                   run1.HasShadow == run2.HasShadow &&
                   IsSameColor(run1.FontColor, run2.FontColor) &&
                   IsSameTextFill(run1.TextFill, run2.TextFill) &&
                   IsSameTextOutline(run1.TextOutline, run2.TextOutline) &&
                   IsSameTextEffects(run1.TextEffects, run2.TextEffects);
        }

        /// <summary>
        /// 判断两个TextFill是否相同
        /// Performs deep equality comparison including gradient stops and pattern fills
        /// </summary>
        private bool IsSameTextFill(TextFillInfo fill1, TextFillInfo fill2)
        {
            // Handle null cases
            if (fill1 == null && fill2 == null) return true;
            if (fill1 == null || fill2 == null) return false;
            
            // Compare basic properties
            if (fill1.HasFill != fill2.HasFill || fill1.FillType != fill2.FillType)
                return false;
            
            // If no fill, they are equal
            if (!fill1.HasFill) return true;
            
            // Compare transparency
            if (Math.Abs(fill1.Transparency - fill2.Transparency) >= 0.01f)
                return false;
            
            // For solid fills, compare color
            if (fill1.FillType == FillType.Solid)
            {
                return IsSameColor(fill1.Color, fill2.Color);
            }
            
            // For gradient fills, perform deep equality comparison
            if (fill1.FillType == FillType.Gradient)
            {
                return IsSameGradient(fill1.Gradient, fill2.Gradient);
            }
            
            // For pattern fills, perform deep equality comparison
            if (fill1.FillType == FillType.Pattern)
            {
                return IsSamePattern(fill1.Pattern, fill2.Pattern);
            }
            
            // Unknown fill type - consider different
            return false;
        }

        /// <summary>
        /// Compare two gradient fills for deep equality
        /// Compares gradient type, angle, and all gradient stops (position and color)
        /// </summary>
        private bool IsSameGradient(GradientInfo grad1, GradientInfo grad2)
        {
            // Handle null cases
            if (grad1 == null && grad2 == null) return true;
            if (grad1 == null || grad2 == null) return false;
            
            // Compare gradient type
            if (grad1.GradientType != grad2.GradientType)
                return false;
            
            // Compare angle (with epsilon for floating point)
            if (Math.Abs(grad1.Angle - grad2.Angle) >= 0.01f)
                return false;
            
            // Compare gradient stops
            if (grad1.Stops == null && grad2.Stops == null) return true;
            if (grad1.Stops == null || grad2.Stops == null) return false;
            
            // Must have same number of stops
            if (grad1.Stops.Count != grad2.Stops.Count)
                return false;
            
            // Compare each gradient stop (position and color)
            for (int i = 0; i < grad1.Stops.Count; i++)
            {
                var stop1 = grad1.Stops[i];
                var stop2 = grad2.Stops[i];
                
                // Compare position (with epsilon for floating point)
                if (Math.Abs(stop1.Position - stop2.Position) >= 0.01f)
                    return false;
                
                // Compare color at this stop
                if (!IsSameColor(stop1.Color, stop2.Color))
                    return false;
            }
            
            return true;
        }

        /// <summary>
        /// Compare two pattern fills for deep equality
        /// Compares pattern type, foreground color, and background color
        /// </summary>
        private bool IsSamePattern(PatternInfo patt1, PatternInfo patt2)
        {
            // Handle null cases
            if (patt1 == null && patt2 == null) return true;
            if (patt1 == null || patt2 == null) return false;
            
            // Compare pattern type
            if (patt1.PatternType != patt2.PatternType)
                return false;
            
            // Compare foreground color
            if (!IsSameColor(patt1.ForegroundColor, patt2.ForegroundColor))
                return false;
            
            // Compare background color
            if (!IsSameColor(patt1.BackgroundColor, patt2.BackgroundColor))
                return false;
            
            return true;
        }

        /// <summary>
        /// 判断两个TextOutline是否相同
        /// </summary>
        private bool IsSameTextOutline(TextOutlineInfo outline1, TextOutlineInfo outline2)
        {
            if (outline1 == null && outline2 == null) return true;
            if (outline1 == null || outline2 == null) return false;
            
            if (outline1.HasOutline != outline2.HasOutline)
                return false;
            
            if (!outline1.HasOutline) return true;
            
            // Compare outline properties
            return Math.Abs(outline1.Width - outline2.Width) < 0.01f &&
                   IsSameColor(outline1.Color, outline2.Color) &&
                   outline1.DashStyle == outline2.DashStyle &&
                   outline1.CompoundLineType == outline2.CompoundLineType &&
                   outline1.CapType == outline2.CapType &&
                   outline1.JoinType == outline2.JoinType &&
                   Math.Abs(outline1.Transparency - outline2.Transparency) < 0.01f;
        }

        /// <summary>
        /// 判断两个TextEffects是否相同
        /// Performs deep equality comparison of all text effects including shadow, glow, reflection, and soft edge
        /// </summary>
        private bool IsSameTextEffects(TextEffectsInfo effects1, TextEffectsInfo effects2)
        {
            // Handle null cases
            if (effects1 == null && effects2 == null) return true;
            if (effects1 == null || effects2 == null) return false;
            
            // Compare HasEffects flag
            if (effects1.HasEffects != effects2.HasEffects)
                return false;
            
            // If neither has effects, they are the same
            if (!effects1.HasEffects) return true;
            
            // Compare individual effect flags
            if (effects1.HasShadow != effects2.HasShadow ||
                effects1.HasGlow != effects2.HasGlow ||
                effects1.HasReflection != effects2.HasReflection ||
                effects1.HasSoftEdge != effects2.HasSoftEdge)
                return false;
            
            // Deep comparison of shadow effect
            if (effects1.HasShadow)
            {
                if (!IsSameShadow(effects1.Shadow, effects2.Shadow))
                    return false;
            }
            
            // Deep comparison of glow effect
            if (effects1.HasGlow)
            {
                if (!IsSameGlow(effects1.Glow, effects2.Glow))
                    return false;
            }
            
            // Deep comparison of reflection effect
            if (effects1.HasReflection)
            {
                if (!IsSameReflection(effects1.Reflection, effects2.Reflection))
                    return false;
            }
            
            // Deep comparison of soft edge effect
            if (effects1.HasSoftEdge)
            {
                if (!IsSameSoftEdge(effects1.SoftEdgeRadius, effects2.SoftEdgeRadius))
                    return false;
            }
            
            return true;
        }

        /// <summary>
        /// Compare two shadow effects for equality
        /// </summary>
        private bool IsSameShadow(ShadowInfo shadow1, ShadowInfo shadow2)
        {
            // Handle null cases
            if (shadow1 == null && shadow2 == null) return true;
            if (shadow1 == null || shadow2 == null) return false;
            
            // Compare shadow type
            if (shadow1.Type != shadow2.Type)
                return false;
            
            // Compare shadow color
            if (!IsSameColor(shadow1.Color, shadow2.Color))
                return false;
            
            // Compare floating-point values with epsilon tolerance
            const float epsilon = 0.01f;
            if (Math.Abs(shadow1.Blur - shadow2.Blur) >= epsilon)
                return false;
            if (Math.Abs(shadow1.Distance - shadow2.Distance) >= epsilon)
                return false;
            if (Math.Abs(shadow1.Angle - shadow2.Angle) >= epsilon)
                return false;
            if (Math.Abs(shadow1.Transparency - shadow2.Transparency) >= epsilon)
                return false;
            
            return true;
        }

        /// <summary>
        /// Compare two glow effects for equality
        /// </summary>
        private bool IsSameGlow(GlowInfo glow1, GlowInfo glow2)
        {
            // Handle null cases
            if (glow1 == null && glow2 == null) return true;
            if (glow1 == null || glow2 == null) return false;
            
            // Compare glow color
            if (!IsSameColor(glow1.Color, glow2.Color))
                return false;
            
            // Compare floating-point values with epsilon tolerance
            const float epsilon = 0.01f;
            if (Math.Abs(glow1.Radius - glow2.Radius) >= epsilon)
                return false;
            if (Math.Abs(glow1.Transparency - glow2.Transparency) >= epsilon)
                return false;
            
            return true;
        }

        /// <summary>
        /// Compare two reflection effects for equality
        /// </summary>
        private bool IsSameReflection(ReflectionInfo reflection1, ReflectionInfo reflection2)
        {
            // Handle null cases
            if (reflection1 == null && reflection2 == null) return true;
            if (reflection1 == null || reflection2 == null) return false;
            
            // Compare all reflection parameters with epsilon tolerance
            const float epsilon = 0.01f;
            if (Math.Abs(reflection1.BlurRadius - reflection2.BlurRadius) >= epsilon)
                return false;
            if (Math.Abs(reflection1.StartOpacity - reflection2.StartOpacity) >= epsilon)
                return false;
            if (Math.Abs(reflection1.StartPosition - reflection2.StartPosition) >= epsilon)
                return false;
            if (Math.Abs(reflection1.EndAlpha - reflection2.EndAlpha) >= epsilon)
                return false;
            if (Math.Abs(reflection1.EndPosition - reflection2.EndPosition) >= epsilon)
                return false;
            if (Math.Abs(reflection1.Distance - reflection2.Distance) >= epsilon)
                return false;
            if (Math.Abs(reflection1.Direction - reflection2.Direction) >= epsilon)
                return false;
            if (Math.Abs(reflection1.FadeDirection - reflection2.FadeDirection) >= epsilon)
                return false;
            if (Math.Abs(reflection1.SkewHorizontal - reflection2.SkewHorizontal) >= epsilon)
                return false;
            if (Math.Abs(reflection1.SkewVertical - reflection2.SkewVertical) >= epsilon)
                return false;
            
            return true;
        }

        /// <summary>
        /// Compare two soft edge radii for equality
        /// </summary>
        private bool IsSameSoftEdge(float radius1, float radius2)
        {
            const float epsilon = 0.01f;
            return Math.Abs(radius1 - radius2) < epsilon;
        }

        /// <summary>
        /// 判断两个颜色是否相同
        /// Compares RGB values, theme color information, and color transforms
        /// </summary>
        private bool IsSameColor(ColorInfo color1, ColorInfo color2)
        {
            // Handle null cases
            if (color1 == null && color2 == null) return true;
            if (color1 == null || color2 == null) return false;
            
            // Compare RGB values
            if (color1.Red != color2.Red || color1.Green != color2.Green || color1.Blue != color2.Blue)
                return false;
            
            // Compare theme color properties
            if (color1.IsThemeColor != color2.IsThemeColor)
                return false;
            
            // If both are theme colors, compare scheme color name
            if (color1.IsThemeColor && color2.IsThemeColor)
            {
                if (color1.SchemeColorName != color2.SchemeColorName)
                    return false;
            }
            
            // Compare color transforms
            return IsSameColorTransforms(color1.Transforms, color2.Transforms);
        }
        
        /// <summary>
        /// 判断两个颜色变换是否相同
        /// Compares all color transform properties (lumMod, lumOff, tint, shade, satMod, satOff, alpha)
        /// </summary>
        private bool IsSameColorTransforms(ColorTransforms t1, ColorTransforms t2)
        {
            // Handle null cases
            if (t1 == null && t2 == null) return true;
            if (t1 == null || t2 == null) return false;
            
            // Compare all transform properties
            return t1.LumMod == t2.LumMod &&
                   t1.LumOff == t2.LumOff &&
                   t1.Tint == t2.Tint &&
                   t1.Shade == t2.Shade &&
                   t1.SatMod == t2.SatMod &&
                   t1.SatOff == t2.SatOff &&
                   t1.Alpha == t2.Alpha;
        }

        /// <summary>
        /// 添加阴影JSON
        /// </summary>
        private void AppendShadowJson(StringBuilder sb, ShadowInfo shadow, bool hasShadow)
        {
            sb.Append("{");
            sb.Append($"\"has_shadow\":{(hasShadow ? 1 : 0)},");
            if (shadow != null && hasShadow)
            {
                sb.Append($"\"color\":\"{shadow.Color?.ToString() ?? "RGB(0, 0, 0)"}\",");
                sb.Append($"\"opacity\":{100 - shadow.Transparency:F1},");
                sb.Append($"\"blur\":{shadow.Blur:F1},");
                sb.Append($"\"offset_x\":{(shadow.Distance * Math.Cos(shadow.Angle * Math.PI / 180)):F1},");
                sb.Append($"\"offset_y\":{(shadow.Distance * Math.Sin(shadow.Angle * Math.PI / 180)):F1},");
                sb.Append($"\"size\":{shadow.Distance:F1},");
                sb.Append($"\"transparency\":{shadow.Transparency:F1},");
                sb.Append($"\"type\":{(int)shadow.Type},");
                sb.Append($"\"style\":0,");
                sb.Append($"\"shadow_type\":\"{GetShadowTypeName(shadow.Type)}\"");
            }
            else
            {
                sb.Append("\"color\":\"RGB(0, 0, 0)\",");
                sb.Append("\"opacity\":0.0,");
                sb.Append("\"blur\":0.0,");
                sb.Append("\"offset_x\":0.0,");
                sb.Append("\"offset_y\":0.0,");
                sb.Append("\"size\":0.0,");
                sb.Append("\"transparency\":0.0,");
                sb.Append("\"type\":0,");
                sb.Append("\"style\":0,");
                sb.Append("\"shadow_type\":\"\"");
            }
            sb.Append("}");
        }

        /// <summary>
        /// 获取阴影类型名称
        /// </summary>
        private string GetShadowTypeName(ShadowType type)
        {
            switch (type)
            {
                case ShadowType.Outer: return "outer";
                case ShadowType.Inner: return "inner";
                case ShadowType.Perspective: return "perspective";
                default: return "";
            }
        }

        /// <summary>
        /// Serialize text fill information to JSON string
        /// Handles solid fill, gradient fill, pattern fill, and no fill
        /// Preserves theme color references and transformations
        /// </summary>
        private string SerializeTextFillToJson(TextFillInfo textFill)
        {
            if (textFill == null)
            {
                return "{\"has_fill\":0,\"fill_type\":\"none\"}";
            }

            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append($"\"has_fill\":{(textFill.HasFill ? 1 : 0)},");
            sb.Append($"\"fill_type\":\"{GetFillTypeName(textFill.FillType)}\"");

            if (textFill.HasFill)
            {
                switch (textFill.FillType)
                {
                    case FillType.Solid:
                        // Serialize solid fill with color and transparency
                        if (textFill.Color != null)
                        {
                            sb.Append($",\"color\":\"{textFill.Color}\"");
                            
                            // Add theme color information if present
                            if (textFill.Color.IsThemeColor && !string.IsNullOrEmpty(textFill.Color.SchemeColorName))
                            {
                                sb.Append($",\"schemeColor\":\"{textFill.Color.SchemeColorName}\"");
                                
                                // Add color transformations if present
                                if (textFill.Color.Transforms != null && textFill.Color.Transforms.HasTransforms)
                                {
                                    sb.Append(",\"colorTransforms\":{");
                                    var transformParts = new List<string>();
                                    if (textFill.Color.Transforms.LumMod.HasValue)
                                        transformParts.Add($"\"lumMod\":{textFill.Color.Transforms.LumMod.Value}");
                                    if (textFill.Color.Transforms.LumOff.HasValue)
                                        transformParts.Add($"\"lumOff\":{textFill.Color.Transforms.LumOff.Value}");
                                    if (textFill.Color.Transforms.Tint.HasValue)
                                        transformParts.Add($"\"tint\":{textFill.Color.Transforms.Tint.Value}");
                                    if (textFill.Color.Transforms.Shade.HasValue)
                                        transformParts.Add($"\"shade\":{textFill.Color.Transforms.Shade.Value}");
                                    if (textFill.Color.Transforms.SatMod.HasValue)
                                        transformParts.Add($"\"satMod\":{textFill.Color.Transforms.SatMod.Value}");
                                    if (textFill.Color.Transforms.SatOff.HasValue)
                                        transformParts.Add($"\"satOff\":{textFill.Color.Transforms.SatOff.Value}");
                                    if (textFill.Color.Transforms.Alpha.HasValue)
                                        transformParts.Add($"\"alpha\":{textFill.Color.Transforms.Alpha.Value}");
                                    sb.Append(string.Join(",", transformParts));
                                    sb.Append("}");
                                }
                            }
                            else if (!string.IsNullOrEmpty(textFill.Color.OriginalHex))
                            {
                                sb.Append($",\"originalHex\":\"{textFill.Color.OriginalHex}\"");
                            }
                        }
                        sb.Append($",\"transparency\":{textFill.Transparency:F2}");
                        break;

                    case FillType.Gradient:
                        // Serialize gradient fill with stops and direction
                        if (textFill.Gradient != null)
                        {
                            sb.Append($",\"gradient_type\":\"{textFill.Gradient.GradientType}\"");
                            sb.Append($",\"angle\":{textFill.Gradient.Angle:F1}");
                            
                            if (textFill.Gradient.Stops != null && textFill.Gradient.Stops.Count > 0)
                            {
                                sb.Append(",\"stops\":[");
                                bool firstStop = true;
                                foreach (var stop in textFill.Gradient.Stops)
                                {
                                    if (!firstStop) sb.Append(",");
                                    sb.Append("{");
                                    sb.Append($"\"position\":{stop.Position:F2}");
                                    if (stop.Color != null)
                                    {
                                        sb.Append($",\"color\":\"{stop.Color}\"");
                                        
                                        // Add theme color information for gradient stops
                                        if (stop.Color.IsThemeColor && !string.IsNullOrEmpty(stop.Color.SchemeColorName))
                                        {
                                            sb.Append($",\"schemeColor\":\"{stop.Color.SchemeColorName}\"");
                                            
                                            if (stop.Color.Transforms != null && stop.Color.Transforms.HasTransforms)
                                            {
                                                sb.Append(",\"colorTransforms\":{");
                                                var transformParts = new List<string>();
                                                if (stop.Color.Transforms.LumMod.HasValue)
                                                    transformParts.Add($"\"lumMod\":{stop.Color.Transforms.LumMod.Value}");
                                                if (stop.Color.Transforms.LumOff.HasValue)
                                                    transformParts.Add($"\"lumOff\":{stop.Color.Transforms.LumOff.Value}");
                                                if (stop.Color.Transforms.Tint.HasValue)
                                                    transformParts.Add($"\"tint\":{stop.Color.Transforms.Tint.Value}");
                                                if (stop.Color.Transforms.Shade.HasValue)
                                                    transformParts.Add($"\"shade\":{stop.Color.Transforms.Shade.Value}");
                                                if (stop.Color.Transforms.SatMod.HasValue)
                                                    transformParts.Add($"\"satMod\":{stop.Color.Transforms.SatMod.Value}");
                                                if (stop.Color.Transforms.SatOff.HasValue)
                                                    transformParts.Add($"\"satOff\":{stop.Color.Transforms.SatOff.Value}");
                                                if (stop.Color.Transforms.Alpha.HasValue)
                                                    transformParts.Add($"\"alpha\":{stop.Color.Transforms.Alpha.Value}");
                                                sb.Append(string.Join(",", transformParts));
                                                sb.Append("}");
                                            }
                                        }
                                    }
                                    sb.Append("}");
                                    firstStop = false;
                                }
                                sb.Append("]");
                            }
                        }
                        break;

                    case FillType.Pattern:
                        // Serialize pattern fill with pattern type and colors
                        if (textFill.Pattern != null)
                        {
                            sb.Append($",\"pattern_type\":\"{EscapeJson(textFill.Pattern.PatternType)}\"");
                            
                            if (textFill.Pattern.ForegroundColor != null)
                            {
                                sb.Append($",\"foreground_color\":\"{textFill.Pattern.ForegroundColor}\"");
                                
                                // Add theme color for foreground
                                if (textFill.Pattern.ForegroundColor.IsThemeColor && 
                                    !string.IsNullOrEmpty(textFill.Pattern.ForegroundColor.SchemeColorName))
                                {
                                    sb.Append($",\"fg_schemeColor\":\"{textFill.Pattern.ForegroundColor.SchemeColorName}\"");
                                    
                                    if (textFill.Pattern.ForegroundColor.Transforms != null && 
                                        textFill.Pattern.ForegroundColor.Transforms.HasTransforms)
                                    {
                                        sb.Append(",\"fg_colorTransforms\":{");
                                        var transformParts = new List<string>();
                                        if (textFill.Pattern.ForegroundColor.Transforms.LumMod.HasValue)
                                            transformParts.Add($"\"lumMod\":{textFill.Pattern.ForegroundColor.Transforms.LumMod.Value}");
                                        if (textFill.Pattern.ForegroundColor.Transforms.LumOff.HasValue)
                                            transformParts.Add($"\"lumOff\":{textFill.Pattern.ForegroundColor.Transforms.LumOff.Value}");
                                        if (textFill.Pattern.ForegroundColor.Transforms.Tint.HasValue)
                                            transformParts.Add($"\"tint\":{textFill.Pattern.ForegroundColor.Transforms.Tint.Value}");
                                        if (textFill.Pattern.ForegroundColor.Transforms.Shade.HasValue)
                                            transformParts.Add($"\"shade\":{textFill.Pattern.ForegroundColor.Transforms.Shade.Value}");
                                        sb.Append(string.Join(",", transformParts));
                                        sb.Append("}");
                                    }
                                }
                            }
                            
                            if (textFill.Pattern.BackgroundColor != null)
                            {
                                sb.Append($",\"background_color\":\"{textFill.Pattern.BackgroundColor}\"");
                                
                                // Add theme color for background
                                if (textFill.Pattern.BackgroundColor.IsThemeColor && 
                                    !string.IsNullOrEmpty(textFill.Pattern.BackgroundColor.SchemeColorName))
                                {
                                    sb.Append($",\"bg_schemeColor\":\"{textFill.Pattern.BackgroundColor.SchemeColorName}\"");
                                    
                                    if (textFill.Pattern.BackgroundColor.Transforms != null && 
                                        textFill.Pattern.BackgroundColor.Transforms.HasTransforms)
                                    {
                                        sb.Append(",\"bg_colorTransforms\":{");
                                        var transformParts = new List<string>();
                                        if (textFill.Pattern.BackgroundColor.Transforms.LumMod.HasValue)
                                            transformParts.Add($"\"lumMod\":{textFill.Pattern.BackgroundColor.Transforms.LumMod.Value}");
                                        if (textFill.Pattern.BackgroundColor.Transforms.LumOff.HasValue)
                                            transformParts.Add($"\"lumOff\":{textFill.Pattern.BackgroundColor.Transforms.LumOff.Value}");
                                        if (textFill.Pattern.BackgroundColor.Transforms.Tint.HasValue)
                                            transformParts.Add($"\"tint\":{textFill.Pattern.BackgroundColor.Transforms.Tint.Value}");
                                        if (textFill.Pattern.BackgroundColor.Transforms.Shade.HasValue)
                                            transformParts.Add($"\"shade\":{textFill.Pattern.BackgroundColor.Transforms.Shade.Value}");
                                        sb.Append(string.Join(",", transformParts));
                                        sb.Append("}");
                                    }
                                }
                            }
                        }
                        break;
                }
            }

            sb.Append("}");
            return sb.ToString();
        }

        /// <summary>
        /// Get fill type name for JSON serialization
        /// </summary>
        private string GetFillTypeName(FillType fillType)
        {
            switch (fillType)
            {
                case FillType.Solid: return "solid";
                case FillType.Gradient: return "gradient";
                case FillType.Pattern: return "pattern";
                case FillType.NoFill: return "none";
                case FillType.Picture: return "picture";
                case FillType.Background: return "background";
                default: return "none";
            }
        }

        /// <summary>
        /// Serialize text outline information to JSON string
        /// Handles outline width, color, dash style, compound type, cap type, and join type
        /// Preserves theme color references and transformations
        /// </summary>
        private string SerializeTextOutlineToJson(TextOutlineInfo textOutline)
        {
            if (textOutline == null)
            {
                return "{\"has_outline\":0}";
            }

            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append($"\"has_outline\":{(textOutline.HasOutline ? 1 : 0)}");

            if (textOutline.HasOutline)
            {
                // Serialize outline width
                sb.Append($",\"width\":{textOutline.Width:F2}");

                // Serialize outline color with theme color preservation
                if (textOutline.Color != null)
                {
                    sb.Append($",\"color\":\"{textOutline.Color}\"");
                    
                    // Add theme color information if present
                    if (textOutline.Color.IsThemeColor && !string.IsNullOrEmpty(textOutline.Color.SchemeColorName))
                    {
                        sb.Append($",\"schemeColor\":\"{textOutline.Color.SchemeColorName}\"");
                        
                        // Add color transformations if present
                        if (textOutline.Color.Transforms != null && textOutline.Color.Transforms.HasTransforms)
                        {
                            sb.Append(",\"colorTransforms\":{");
                            var transformParts = new List<string>();
                            if (textOutline.Color.Transforms.LumMod.HasValue)
                                transformParts.Add($"\"lumMod\":{textOutline.Color.Transforms.LumMod.Value}");
                            if (textOutline.Color.Transforms.LumOff.HasValue)
                                transformParts.Add($"\"lumOff\":{textOutline.Color.Transforms.LumOff.Value}");
                            if (textOutline.Color.Transforms.Tint.HasValue)
                                transformParts.Add($"\"tint\":{textOutline.Color.Transforms.Tint.Value}");
                            if (textOutline.Color.Transforms.Shade.HasValue)
                                transformParts.Add($"\"shade\":{textOutline.Color.Transforms.Shade.Value}");
                            if (textOutline.Color.Transforms.SatMod.HasValue)
                                transformParts.Add($"\"satMod\":{textOutline.Color.Transforms.SatMod.Value}");
                            if (textOutline.Color.Transforms.SatOff.HasValue)
                                transformParts.Add($"\"satOff\":{textOutline.Color.Transforms.SatOff.Value}");
                            if (textOutline.Color.Transforms.Alpha.HasValue)
                                transformParts.Add($"\"alpha\":{textOutline.Color.Transforms.Alpha.Value}");
                            sb.Append(string.Join(",", transformParts));
                            sb.Append("}");
                        }
                    }
                    else if (!string.IsNullOrEmpty(textOutline.Color.OriginalHex))
                    {
                        sb.Append($",\"originalHex\":\"{textOutline.Color.OriginalHex}\"");
                    }
                }

                // Serialize dash style
                sb.Append($",\"dash_style\":\"{GetDashStyleName(textOutline.DashStyle)}\"");

                // Serialize compound line type
                if (!string.IsNullOrEmpty(textOutline.CompoundLineType))
                {
                    sb.Append($",\"compound_type\":\"{textOutline.CompoundLineType.ToLower()}\"");
                }

                // Serialize cap type
                if (!string.IsNullOrEmpty(textOutline.CapType))
                {
                    sb.Append($",\"cap_type\":\"{textOutline.CapType.ToLower()}\"");
                }

                // Serialize join type
                if (!string.IsNullOrEmpty(textOutline.JoinType))
                {
                    sb.Append($",\"join_type\":\"{textOutline.JoinType.ToLower()}\"");
                }

                // Serialize transparency
                sb.Append($",\"transparency\":{textOutline.Transparency:F2}");
            }

            sb.Append("}");
            return sb.ToString();
        }

        /// <summary>
        /// Get dash style name for JSON serialization
        /// </summary>
        private string GetDashStyleName(LineDashStyle dashStyle)
        {
            switch (dashStyle)
            {
                case LineDashStyle.Solid: return "solid";
                case LineDashStyle.SquareDot: return "square_dot";
                case LineDashStyle.RoundDot: return "round_dot";
                case LineDashStyle.Dash: return "dash";
                case LineDashStyle.DashDot: return "dash_dot";
                case LineDashStyle.LongDash: return "long_dash";
                case LineDashStyle.LongDashDot: return "long_dash_dot";
                case LineDashStyle.LongDashDotDot: return "long_dash_dot_dot";
                default: return "solid";
            }
        }

        /// <summary>
        /// Serialize text effects information to JSON string
        /// Handles shadow, glow, reflection, and soft edges
        /// Supports multiple effects on the same text run
        /// </summary>
        private string SerializeTextEffectsToJson(TextEffectsInfo textEffects)
        {
            if (textEffects == null)
            {
                return "{\"has_effects\":0}";
            }

            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append($"\"has_effects\":{(textEffects.HasEffects ? 1 : 0)}");

            if (textEffects.HasEffects)
            {
                // Serialize shadow effect
                sb.Append(",\"shadow\":{");
                sb.Append($"\"has_shadow\":{(textEffects.HasShadow ? 1 : 0)}");
                
                if (textEffects.HasShadow && textEffects.Shadow != null)
                {
                    sb.Append($",\"type\":\"{GetShadowTypeName(textEffects.Shadow.Type)}\"");
                    
                    if (textEffects.Shadow.Color != null)
                    {
                        sb.Append($",\"color\":\"{textEffects.Shadow.Color}\"");
                        
                        // Add theme color information if present
                        if (textEffects.Shadow.Color.IsThemeColor && !string.IsNullOrEmpty(textEffects.Shadow.Color.SchemeColorName))
                        {
                            sb.Append($",\"schemeColor\":\"{textEffects.Shadow.Color.SchemeColorName}\"");
                            
                            if (textEffects.Shadow.Color.Transforms != null && textEffects.Shadow.Color.Transforms.HasTransforms)
                            {
                                sb.Append(",\"colorTransforms\":{");
                                var transformParts = new List<string>();
                                if (textEffects.Shadow.Color.Transforms.LumMod.HasValue)
                                    transformParts.Add($"\"lumMod\":{textEffects.Shadow.Color.Transforms.LumMod.Value}");
                                if (textEffects.Shadow.Color.Transforms.LumOff.HasValue)
                                    transformParts.Add($"\"lumOff\":{textEffects.Shadow.Color.Transforms.LumOff.Value}");
                                if (textEffects.Shadow.Color.Transforms.Tint.HasValue)
                                    transformParts.Add($"\"tint\":{textEffects.Shadow.Color.Transforms.Tint.Value}");
                                if (textEffects.Shadow.Color.Transforms.Shade.HasValue)
                                    transformParts.Add($"\"shade\":{textEffects.Shadow.Color.Transforms.Shade.Value}");
                                if (textEffects.Shadow.Color.Transforms.Alpha.HasValue)
                                    transformParts.Add($"\"alpha\":{textEffects.Shadow.Color.Transforms.Alpha.Value}");
                                sb.Append(string.Join(",", transformParts));
                                sb.Append("}");
                            }
                        }
                    }
                    
                    sb.Append($",\"blur\":{textEffects.Shadow.Blur:F2}");
                    sb.Append($",\"distance\":{textEffects.Shadow.Distance:F2}");
                    sb.Append($",\"angle\":{textEffects.Shadow.Angle:F2}");
                    sb.Append($",\"transparency\":{textEffects.Shadow.Transparency:F2}");
                }
                
                sb.Append("}");

                // Serialize glow effect
                sb.Append(",\"glow\":{");
                sb.Append($"\"has_glow\":{(textEffects.HasGlow ? 1 : 0)}");
                
                if (textEffects.HasGlow && textEffects.Glow != null)
                {
                    sb.Append($",\"radius\":{textEffects.Glow.Radius:F2}");
                    
                    if (textEffects.Glow.Color != null)
                    {
                        sb.Append($",\"color\":\"{textEffects.Glow.Color}\"");
                        
                        // Add theme color information if present
                        if (textEffects.Glow.Color.IsThemeColor && !string.IsNullOrEmpty(textEffects.Glow.Color.SchemeColorName))
                        {
                            sb.Append($",\"schemeColor\":\"{textEffects.Glow.Color.SchemeColorName}\"");
                            
                            if (textEffects.Glow.Color.Transforms != null && textEffects.Glow.Color.Transforms.HasTransforms)
                            {
                                sb.Append(",\"colorTransforms\":{");
                                var transformParts = new List<string>();
                                if (textEffects.Glow.Color.Transforms.LumMod.HasValue)
                                    transformParts.Add($"\"lumMod\":{textEffects.Glow.Color.Transforms.LumMod.Value}");
                                if (textEffects.Glow.Color.Transforms.LumOff.HasValue)
                                    transformParts.Add($"\"lumOff\":{textEffects.Glow.Color.Transforms.LumOff.Value}");
                                if (textEffects.Glow.Color.Transforms.Tint.HasValue)
                                    transformParts.Add($"\"tint\":{textEffects.Glow.Color.Transforms.Tint.Value}");
                                if (textEffects.Glow.Color.Transforms.Shade.HasValue)
                                    transformParts.Add($"\"shade\":{textEffects.Glow.Color.Transforms.Shade.Value}");
                                if (textEffects.Glow.Color.Transforms.Alpha.HasValue)
                                    transformParts.Add($"\"alpha\":{textEffects.Glow.Color.Transforms.Alpha.Value}");
                                sb.Append(string.Join(",", transformParts));
                                sb.Append("}");
                            }
                        }
                    }
                    
                    sb.Append($",\"transparency\":{textEffects.Glow.Transparency:F2}");
                }
                
                sb.Append("}");

                // Serialize reflection effect
                sb.Append(",\"reflection\":{");
                sb.Append($"\"has_reflection\":{(textEffects.HasReflection ? 1 : 0)}");
                
                if (textEffects.HasReflection && textEffects.Reflection != null)
                {
                    sb.Append($",\"blur_radius\":{textEffects.Reflection.BlurRadius:F2}");
                    sb.Append($",\"start_opacity\":{textEffects.Reflection.StartOpacity:F2}");
                    sb.Append($",\"start_position\":{textEffects.Reflection.StartPosition:F2}");
                    sb.Append($",\"end_alpha\":{textEffects.Reflection.EndAlpha:F2}");
                    sb.Append($",\"end_position\":{textEffects.Reflection.EndPosition:F2}");
                    sb.Append($",\"distance\":{textEffects.Reflection.Distance:F2}");
                    sb.Append($",\"direction\":{textEffects.Reflection.Direction:F2}");
                    sb.Append($",\"fade_direction\":{textEffects.Reflection.FadeDirection:F2}");
                    sb.Append($",\"skew_horizontal\":{textEffects.Reflection.SkewHorizontal:F2}");
                    sb.Append($",\"skew_vertical\":{textEffects.Reflection.SkewVertical:F2}");
                }
                
                sb.Append("}");

                // Serialize soft edge effect
                sb.Append(",\"soft_edge\":{");
                sb.Append($"\"has_soft_edge\":{(textEffects.HasSoftEdge ? 1 : 0)}");
                
                if (textEffects.HasSoftEdge)
                {
                    sb.Append($",\"radius\":{textEffects.SoftEdgeRadius:F2}");
                }
                
                sb.Append("}");
            }

            sb.Append("}");
            return sb.ToString();
        }
        
        private string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
        }

        public void FromJson(object jsonData)
        {
            try
            {
                if (jsonData == null) { HasText = false; TextContent = ""; return; }
                dynamic data = jsonData;
                
                try { HasText = Convert.ToInt32(data.hasText) != 0; } catch { HasText = false; }
                try { TextContent = data.content?.ToString() ?? ""; } catch { TextContent = ""; }
                try { FontName = data.fontName?.ToString() ?? ""; } catch { FontName = ""; }
                try { FontSize = Convert.ToSingle(data.fontSize); } catch { FontSize = 12; }
                try { FontColor = ColorInfo.Parse(data.fontColor?.ToString()); } catch { FontColor = new ColorInfo(); }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从JSON加载文本信息时出错: {ex.Message}");
                HasText = false;
                TextContent = "";
            }
        }
    }
}
