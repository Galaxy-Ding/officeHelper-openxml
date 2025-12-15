using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Core.Repair
{
    /// <summary>
    /// 结构验证器 - 负责修复PPTX XML结构完整性问题
    /// 处理P0级别的关键修复：缺失必需元素和属性
    /// </summary>
    public class StructureValidator
    {
        /// <summary>
        /// 验证并修复幻灯片结构
        /// </summary>
        /// <param name="slide">要修复的幻灯片</param>
        public void ValidateAndFixSlide(Slide slide)
        {
            if (slide == null)
                throw new ArgumentNullException(nameof(slide));

            EnsureRequiredElements(slide);
            FixMissingAttributes(slide);
            ValidateStructure(slide);
        }

        /// <summary>
        /// 验证并修复母版幻灯片结构
        /// </summary>
        /// <param name="master">要修复的母版幻灯片</param>
        public void ValidateAndFixMaster(SlideMaster master)
        {
            if (master == null)
                throw new ArgumentNullException(nameof(master));

            EnsureRequiredElements(master);
            FixMissingAttributes(master);
            ValidateStructure(master);
        }

        /// <summary>
        /// 验证并修复主题结构
        /// </summary>
        /// <param name="theme">要修复的主题</param>
        public void ValidateAndFixTheme(Theme theme)
        {
            if (theme == null)
                throw new ArgumentNullException(nameof(theme));

            EnsureRequiredElements(theme);
            ValidateStructure(theme);
        }

        /// <summary>
        /// 确保必需元素存在
        /// </summary>
        /// <param name="element">要检查的OpenXML元素</param>
        public void EnsureRequiredElements(OpenXmlElement element)
        {
            switch (element)
            {
                case Slide slide:
                    EnsureSlideRequiredElements(slide);
                    break;
                case SlideMaster master:
                    EnsureMasterRequiredElements(master);
                    break;
                case Theme theme:
                    EnsureThemeRequiredElements(theme);
                    break;
            }
        }

        /// <summary>
        /// 修复缺失属性
        /// </summary>
        /// <param name="element">要修复的OpenXML元素</param>
        public void FixMissingAttributes(OpenXmlElement element)
        {
            switch (element)
            {
                case Slide slide:
                    FixSlideMissingAttributes(slide);
                    break;
                case SlideMaster master:
                    FixMasterMissingAttributes(master);
                    break;
            }
        }

        /// <summary>
        /// 验证结构完整性
        /// </summary>
        /// <param name="element">要验证的OpenXML元素</param>
        /// <returns>结构是否有效</returns>
        public bool ValidateStructure(OpenXmlElement element)
        {
            try
            {
                // 基本结构验证
                if (element == null)
                    return false;

                // 检查是否有必需的子元素
                switch (element)
                {
                    case Slide slide:
                        return ValidateSlideStructure(slide);
                    case SlideMaster master:
                        return ValidateMasterStructure(master);
                    case Theme theme:
                        return ValidateThemeStructure(theme);
                    default:
                        return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        #region 幻灯片修复方法

        /// <summary>
        /// 确保幻灯片必需元素
        /// </summary>
        private void EnsureSlideRequiredElements(Slide slide)
        {
            // 补充 effectLst - 效果列表
            var shapes = slide.Descendants<A.Shape>().ToList();
            foreach (var shape in shapes)
            {
                var shapeProps = shape.ShapeProperties;
                if (shapeProps != null && !shapeProps.Elements<A.EffectList>().Any())
                {
                    shapeProps.Append(new A.EffectList());
                }
            }

            // 补充 spAutoFit - 文本框自动适应
            var textBodies = slide.Descendants<A.TextBody>().ToList();
            foreach (var textBody in textBodies)
            {
                var bodyProps = textBody.BodyProperties;
                if (bodyProps != null && !bodyProps.Elements<A.ShapeAutoFit>().Any())
                {
                    bodyProps.Append(new A.ShapeAutoFit());
                }
            }

            // 补充 timing - 动画时间
            if (!slide.Elements<Timing>().Any())
            {
                slide.Append(new Timing());
            }

            // 补充 extLst - 扩展列表
            EnsureExtensionList(slide);
        }

        /// <summary>
        /// 修复幻灯片缺失属性
        /// </summary>
        private void FixSlideMissingAttributes(Slide slide)
        {
            // 修复文本运行属性
            var textRuns = slide.Descendants<A.TextRun>().ToList();
            foreach (var run in textRuns)
            {
                var runProps = run.RunProperties;
                if (runProps == null)
                {
                    runProps = new A.RunProperties();
                    run.InsertAt(runProps, 0);
                }

                // 确保必需属性存在
                if (runProps.Language == null)
                    runProps.Language = "zh-CN";
                
                if (runProps.AlternativeLanguage == null)
                    runProps.AlternativeLanguage = "en-US";
                
                if (!runProps.Dirty.HasValue)
                    runProps.Dirty = false;
            }

            // 修复形状属性
            var shapes = slide.Descendants<A.Shape>().ToList();
            foreach (var shape in shapes)
            {
                var nonVisualProps = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
                if (nonVisualProps != null)
                {
                    // 确保形状ID存在
                    if (!nonVisualProps.Id.HasValue)
                        nonVisualProps.Id = 1u;
                    
                    // 确保形状名称存在
                    if (string.IsNullOrEmpty(nonVisualProps.Name))
                        nonVisualProps.Name = $"Shape_{nonVisualProps.Id}";
                }
            }
        }

        /// <summary>
        /// 验证幻灯片结构
        /// </summary>
        private bool ValidateSlideStructure(Slide slide)
        {
            // 检查基本结构
            if (slide.CommonSlideData?.ShapeTree == null)
                return false;

            // 检查是否有形状
            var shapes = slide.Descendants<A.Shape>().ToList();
            if (shapes.Count == 0)
                return true; // 空幻灯片是有效的

            // 检查每个形状的基本结构
            foreach (var shape in shapes)
            {
                if (shape.NonVisualShapeProperties == null ||
                    shape.ShapeProperties == null)
                    return false;
            }

            return true;
        }

        #endregion

        #region 母版幻灯片修复方法

        /// <summary>
        /// 确保母版幻灯片必需元素
        /// </summary>
        private void EnsureMasterRequiredElements(SlideMaster master)
        {
            // 补充 effectLst
            var shapes = master.Descendants<A.Shape>().ToList();
            foreach (var shape in shapes)
            {
                var shapeProps = shape.ShapeProperties;
                if (shapeProps != null && !shapeProps.Elements<A.EffectList>().Any())
                {
                    shapeProps.Append(new A.EffectList());
                }
            }

            // 补充 extLst
            EnsureExtensionList(master);
        }

        /// <summary>
        /// 修复母版幻灯片缺失属性
        /// </summary>
        private void FixMasterMissingAttributes(SlideMaster master)
        {
            // 修复文本运行属性
            var textRuns = master.Descendants<A.TextRun>().ToList();
            foreach (var run in textRuns)
            {
                var runProps = run.RunProperties;
                if (runProps == null)
                {
                    runProps = new A.RunProperties();
                    run.InsertAt(runProps, 0);
                }

                if (runProps.Language == null)
                    runProps.Language = "zh-CN";
                
                if (runProps.AlternativeLanguage == null)
                    runProps.AlternativeLanguage = "en-US";
                
                if (!runProps.Dirty.HasValue)
                    runProps.Dirty = false;
            }
        }

        /// <summary>
        /// 验证母版幻灯片结构
        /// </summary>
        private bool ValidateMasterStructure(SlideMaster master)
        {
            // 检查基本结构
            if (master.CommonSlideData?.ShapeTree == null)
                return false;

            // 检查文本样式
            if (master.TextStyles == null)
                return false;

            return true;
        }

        #endregion

        #region 主题修复方法

        /// <summary>
        /// 确保主题必需元素
        /// </summary>
        private void EnsureThemeRequiredElements(Theme theme)
        {
            // 确保主题元素存在
            if (theme.ThemeElements == null)
            {
                theme.ThemeElements = new A.ThemeElements();
            }

            // 确保颜色方案存在
            var colorScheme = theme.ThemeElements.Elements<A.ColorScheme>().FirstOrDefault();
            if (colorScheme == null)
            {
                colorScheme = CreateDefaultColorScheme();
                theme.ThemeElements.Append(colorScheme);
            }

            // 确保字体方案存在
            var fontScheme = theme.ThemeElements.Elements<A.FontScheme>().FirstOrDefault();
            if (fontScheme == null)
            {
                fontScheme = CreateDefaultFontScheme();
                theme.ThemeElements.Append(fontScheme);
            }

            // 确保格式方案存在
            var formatScheme = theme.ThemeElements.Elements<A.FormatScheme>().FirstOrDefault();
            if (formatScheme == null)
            {
                formatScheme = CreateDefaultFormatScheme();
                theme.ThemeElements.Append(formatScheme);
            }
        }

        /// <summary>
        /// 验证主题结构
        /// </summary>
        private bool ValidateThemeStructure(Theme theme)
        {
            // 检查主题元素
            if (theme.ThemeElements == null)
                return false;

            // 检查颜色方案
            if (!theme.ThemeElements.Elements<A.ColorScheme>().Any())
                return false;

            // 检查字体方案
            if (!theme.ThemeElements.Elements<A.FontScheme>().Any())
                return false;

            // 检查格式方案
            if (!theme.ThemeElements.Elements<A.FormatScheme>().Any())
                return false;

            return true;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 确保扩展列表存在
        /// </summary>
        private void EnsureExtensionList(OpenXmlElement element)
        {
            if (!element.Elements<A.ExtensionList>().Any())
            {
                element.Append(new A.ExtensionList());
            }
        }

        /// <summary>
        /// 创建默认颜色方案
        /// </summary>
        private A.ColorScheme CreateDefaultColorScheme()
        {
            return new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex() { Val = "44546A" }),
                new A.Light2Color(new A.RgbColorModelHex() { Val = "E7E6E6" }),
                new A.Accent1Color(new A.RgbColorModelHex() { Val = "4472C4" }),
                new A.Accent2Color(new A.RgbColorModelHex() { Val = "ED7D31" }),
                new A.Accent3Color(new A.RgbColorModelHex() { Val = "A5A5A5" }),
                new A.Accent4Color(new A.RgbColorModelHex() { Val = "FFC000" }),
                new A.Accent5Color(new A.RgbColorModelHex() { Val = "5B9BD5" }),
                new A.Accent6Color(new A.RgbColorModelHex() { Val = "70AD47" }),
                new A.Hyperlink(new A.RgbColorModelHex() { Val = "0563C1" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "954F72" })
            ) { Name = "Office" };
        }

        /// <summary>
        /// 创建默认字体方案
        /// </summary>
        private A.FontScheme CreateDefaultFontScheme()
        {
            return new A.FontScheme(
                new A.MajorFont(
                    new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" },
                    new A.EastAsianFont() { Typeface = "" },
                    new A.ComplexScriptFont() { Typeface = "" }
                ),
                new A.MinorFont(
                    new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" },
                    new A.EastAsianFont() { Typeface = "" },
                    new A.ComplexScriptFont() { Typeface = "" }
                )
            ) { Name = "Office" };
        }

        /// <summary>
        /// 创建默认格式方案
        /// </summary>
        private A.FormatScheme CreateDefaultFormatScheme()
        {
            return new A.FormatScheme(
                new A.FillStyleList(
                    new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor })
                ),
                new A.LineStyleList(
                    new A.Outline(
                        new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
                        new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
                    ) { Width = 9525 }
                ),
                new A.EffectStyleList(
                    new A.EffectStyle(
                        new A.EffectList(
                            new A.OuterShadow(
                                new A.RgbColorModelHex(new A.Alpha() { Val = 38000 }) { Val = "000000" }
                            ) { BlurRadius = 40000, Distance = 20000, Direction = 5400000, RotateWithShape = false }
                        )
                    )
                )
            ) { Name = "Office" };
        }

        #endregion
    }
}
