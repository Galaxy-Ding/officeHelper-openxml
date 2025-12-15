using System;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// Text outline information for text runs
    /// Represents outline/border styling applied to text characters
    /// </summary>
    public class TextOutlineInfo
    {
        /// <summary>Whether the text has an outline</summary>
        public bool HasOutline { get; set; }
        
        /// <summary>Width of the outline in points</summary>
        public float Width { get; set; }
        
        /// <summary>Color of the outline</summary>
        public ColorInfo Color { get; set; }
        
        /// <summary>Dash style of the outline</summary>
        public LineDashStyle DashStyle { get; set; }
        
        /// <summary>Compound line type (Single, Double, Thick/Thin, etc.)</summary>
        public string CompoundLineType { get; set; }
        
        /// <summary>Cap type (Flat, Round, Square)</summary>
        public string CapType { get; set; }
        
        /// <summary>Join type (Round, Bevel, Miter)</summary>
        public string JoinType { get; set; }
        
        /// <summary>Transparency level (0.0 = opaque, 1.0 = fully transparent)</summary>
        public float Transparency { get; set; }

        public TextOutlineInfo()
        {
            HasOutline = false;
            Width = 0.0f;
            Color = new ColorInfo();
            DashStyle = LineDashStyle.Solid;
            CompoundLineType = "Single";
            CapType = "Flat";
            JoinType = "Round";
            Transparency = 0.0f;
        }

        public override string ToString()
        {
            if (!HasOutline)
                return "No Outline";
            
            return $"Outline: {Width:F2}pt, {Color}, {DashStyle}";
        }
    }
}
