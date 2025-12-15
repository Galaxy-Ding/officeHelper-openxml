using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// Gradient stop information for gradient fills
    /// </summary>
    public class GradientStop
    {
        /// <summary>Position of the gradient stop (0.0 to 1.0)</summary>
        public float Position { get; set; }
        
        /// <summary>Color at this gradient stop</summary>
        public ColorInfo Color { get; set; }

        public GradientStop()
        {
            Position = 0.0f;
            Color = new ColorInfo();
        }

        public GradientStop(float position, ColorInfo color)
        {
            Position = position;
            Color = color ?? new ColorInfo();
        }
    }

    /// <summary>
    /// Gradient fill information
    /// </summary>
    public class GradientInfo
    {
        /// <summary>Gradient type (Linear, Radial, Path)</summary>
        public string GradientType { get; set; }
        
        /// <summary>List of gradient stops</summary>
        public List<GradientStop> Stops { get; set; }
        
        /// <summary>Angle for linear gradients (in degrees)</summary>
        public float Angle { get; set; }

        public GradientInfo()
        {
            GradientType = "Linear";
            Stops = new List<GradientStop>();
            Angle = 0.0f;
        }
    }

    /// <summary>
    /// Pattern fill information
    /// </summary>
    public class PatternInfo
    {
        /// <summary>Pattern type</summary>
        public string PatternType { get; set; }
        
        /// <summary>Foreground color of the pattern</summary>
        public ColorInfo ForegroundColor { get; set; }
        
        /// <summary>Background color of the pattern</summary>
        public ColorInfo BackgroundColor { get; set; }

        public PatternInfo()
        {
            PatternType = "None";
            ForegroundColor = new ColorInfo();
            BackgroundColor = new ColorInfo();
        }
    }

    /// <summary>
    /// Text fill information for text runs
    /// Represents fill styling applied to text characters
    /// </summary>
    public class TextFillInfo
    {
        /// <summary>Whether the text has fill</summary>
        public bool HasFill { get; set; }
        
        /// <summary>Type of fill (Solid, Gradient, Pattern, NoFill)</summary>
        public FillType FillType { get; set; }
        
        /// <summary>Color for solid fill</summary>
        public ColorInfo Color { get; set; }
        
        /// <summary>Transparency level (0.0 = opaque, 1.0 = fully transparent)</summary>
        public float Transparency { get; set; }
        
        /// <summary>Gradient fill information (when FillType is Gradient)</summary>
        public GradientInfo Gradient { get; set; }
        
        /// <summary>Pattern fill information (when FillType is Pattern)</summary>
        public PatternInfo Pattern { get; set; }

        public TextFillInfo()
        {
            HasFill = false;
            FillType = FillType.NoFill;
            Color = new ColorInfo();
            Transparency = 0.0f;
            Gradient = null;
            Pattern = null;
        }

        public override string ToString()
        {
            if (!HasFill)
                return "No Fill";
            
            switch (FillType)
            {
                case FillType.Solid:
                    return $"Solid Fill: {Color}";
                case FillType.Gradient:
                    return $"Gradient Fill: {Gradient?.GradientType ?? "Unknown"}";
                case FillType.Pattern:
                    return $"Pattern Fill: {Pattern?.PatternType ?? "Unknown"}";
                default:
                    return $"Fill Type: {FillType}";
            }
        }
    }
}
