using System;

namespace OfficeHelperOpenXml.Models
{
    /// <summary>
    /// Glow effect information
    /// </summary>
    public class GlowInfo
    {
        /// <summary>Radius of the glow effect in points</summary>
        public float Radius { get; set; }
        
        /// <summary>Color of the glow</summary>
        public ColorInfo Color { get; set; }
        
        /// <summary>Transparency level (0.0 = opaque, 1.0 = fully transparent)</summary>
        public float Transparency { get; set; }

        public GlowInfo()
        {
            Radius = 0.0f;
            Color = new ColorInfo();
            Transparency = 0.0f;
        }

        public override string ToString()
        {
            return $"Glow: {Radius:F2}pt, {Color}";
        }
    }

    /// <summary>
    /// Reflection effect information
    /// </summary>
    public class ReflectionInfo
    {
        /// <summary>Blur radius in points</summary>
        public float BlurRadius { get; set; }
        
        /// <summary>Start opacity (0.0 to 1.0)</summary>
        public float StartOpacity { get; set; }
        
        /// <summary>Start position (0.0 to 1.0)</summary>
        public float StartPosition { get; set; }
        
        /// <summary>End alpha/opacity (0.0 to 1.0)</summary>
        public float EndAlpha { get; set; }
        
        /// <summary>End position (0.0 to 1.0)</summary>
        public float EndPosition { get; set; }
        
        /// <summary>Distance from the object in points</summary>
        public float Distance { get; set; }
        
        /// <summary>Direction angle in degrees</summary>
        public float Direction { get; set; }
        
        /// <summary>Fade direction angle in degrees</summary>
        public float FadeDirection { get; set; }
        
        /// <summary>Horizontal skew angle in degrees</summary>
        public float SkewHorizontal { get; set; }
        
        /// <summary>Vertical skew angle in degrees</summary>
        public float SkewVertical { get; set; }

        public ReflectionInfo()
        {
            BlurRadius = 0.0f;
            StartOpacity = 1.0f;
            StartPosition = 0.0f;
            EndAlpha = 0.0f;
            EndPosition = 1.0f;
            Distance = 0.0f;
            Direction = 0.0f;
            FadeDirection = 0.0f;
            SkewHorizontal = 0.0f;
            SkewVertical = 0.0f;
        }

        public override string ToString()
        {
            return $"Reflection: Blur={BlurRadius:F2}pt, Distance={Distance:F2}pt";
        }
    }

    /// <summary>
    /// Text effects information for text runs
    /// Represents visual effects applied to text characters
    /// </summary>
    public class TextEffectsInfo
    {
        /// <summary>Whether the text has any effects</summary>
        public bool HasEffects { get; set; }
        
        /// <summary>Shadow effect information</summary>
        public ShadowInfo Shadow { get; set; }
        
        /// <summary>Whether the text has a shadow</summary>
        public bool HasShadow { get; set; }
        
        /// <summary>Glow effect information</summary>
        public GlowInfo Glow { get; set; }
        
        /// <summary>Whether the text has a glow effect</summary>
        public bool HasGlow { get; set; }
        
        /// <summary>Reflection effect information</summary>
        public ReflectionInfo Reflection { get; set; }
        
        /// <summary>Whether the text has a reflection effect</summary>
        public bool HasReflection { get; set; }
        
        /// <summary>Soft edge radius in points</summary>
        public float SoftEdgeRadius { get; set; }
        
        /// <summary>Whether the text has soft edges</summary>
        public bool HasSoftEdge { get; set; }

        public TextEffectsInfo()
        {
            HasEffects = false;
            Shadow = null;
            HasShadow = false;
            Glow = null;
            HasGlow = false;
            Reflection = null;
            HasReflection = false;
            SoftEdgeRadius = 0.0f;
            HasSoftEdge = false;
        }

        public override string ToString()
        {
            if (!HasEffects)
                return "No Effects";
            
            var effects = new System.Collections.Generic.List<string>();
            if (HasShadow) effects.Add("Shadow");
            if (HasGlow) effects.Add("Glow");
            if (HasReflection) effects.Add("Reflection");
            if (HasSoftEdge) effects.Add("Soft Edge");
            
            return $"Effects: {string.Join(", ", effects)}";
        }
    }
}
