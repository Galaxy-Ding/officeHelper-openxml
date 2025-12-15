using System;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// 文本对齐方式
    /// </summary>
    public enum TextAlignment
    {
        Left = 1,
        Center = 2,
        Right = 3,
        Justify = 4,
        Distributed = 5
    }

    /// <summary>
    /// 垂直对齐方式
    /// </summary>
    public enum VerticalAlignment
    {
        Top = 1,
        Middle = 2,
        Bottom = 3
    }

    /// <summary>
    /// 文本运行信息（一段具有相同格式的文本）
    /// </summary>
    public class TextRunInfo
    {
        public string Text { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikethrough { get; set; }
        public ColorInfo FontColor { get; set; }
        public ColorInfo HighlightColor { get; set; }
        public bool HasHighlight { get; set; }
        
        /// <summary>
        /// Shadow information (obsolete - use TextEffects.Shadow instead)
        /// </summary>
        [Obsolete("Use TextEffects.Shadow instead")]
        public ShadowInfo Shadow { get; set; }
        
        /// <summary>
        /// Whether the text has shadow (obsolete - use TextEffects.HasShadow instead)
        /// </summary>
        [Obsolete("Use TextEffects.HasShadow instead")]
        public bool HasShadow { get; set; }
        
        public int? Superscript { get; set; }
        public int? Subscript { get; set; }
        public float CharacterSpacing { get; set; }
        
        /// <summary>
        /// Text fill styling (WordArt)
        /// </summary>
        public TextFillInfo TextFill { get; set; }
        
        /// <summary>
        /// Text outline styling (WordArt)
        /// </summary>
        public TextOutlineInfo TextOutline { get; set; }
        
        /// <summary>
        /// Text effects styling (WordArt) - includes shadow, glow, reflection, soft edges
        /// </summary>
        public TextEffectsInfo TextEffects { get; set; }

        public TextRunInfo()
        {
            Text = string.Empty;
            FontName = "宋体";
            FontSize = 12;
            IsBold = false;
            IsItalic = false;
            IsUnderline = false;
            IsStrikethrough = false;
            FontColor = new ColorInfo(0, 0, 0, false);
            HighlightColor = new ColorInfo();
            HasHighlight = false;
            Shadow = new ShadowInfo();
            HasShadow = false;
            Superscript = null;
            Subscript = null;
            CharacterSpacing = 0;
            
            // Initialize WordArt properties
            TextFill = new TextFillInfo();
            TextOutline = new TextOutlineInfo();
            TextEffects = new TextEffectsInfo();
        }

        public override string ToString()
        {
            return $"[{FontName} {FontSize}pt] {Text}";
        }
    }

    /// <summary>
    /// 段落信息
    /// </summary>
    public class ParagraphInfo
    {
        public System.Collections.Generic.List<TextRunInfo> Runs { get; set; }
        public TextAlignment Alignment { get; set; }
        public float LineSpacing { get; set; }
        public float SpaceBefore { get; set; }
        public float SpaceAfter { get; set; }
        public float LeftIndent { get; set; }
        public float RightIndent { get; set; }
        public float FirstLineIndent { get; set; }
        public int BulletType { get; set; }
        public string BulletChar { get; set; }
        public int Level { get; set; }

        public ParagraphInfo()
        {
            Runs = new System.Collections.Generic.List<TextRunInfo>();
            Alignment = TextAlignment.Left;
            LineSpacing = 1.0f;
            SpaceBefore = 0;
            SpaceAfter = 0;
            LeftIndent = 0;
            RightIndent = 0;
            FirstLineIndent = 0;
            BulletType = 0;
            BulletChar = string.Empty;
            Level = 0;
        }

        public string GetPlainText()
        {
            var sb = new System.Text.StringBuilder();
            foreach (var run in Runs)
            {
                sb.Append(run.Text);
            }
            return sb.ToString();
        }
    }
}
