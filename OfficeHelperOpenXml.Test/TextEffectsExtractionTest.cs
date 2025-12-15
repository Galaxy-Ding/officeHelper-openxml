using System;
using Xunit;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Test to verify text effects extraction and model structure
    /// </summary>
    public class TextEffectsExtractionTest
    {
        [Fact]
        public void TestTextEffectsInfoInitialization()
        {
            // Create a TextEffectsInfo instance
            var textEffects = new TextEffectsInfo();
            
            // Verify default values
            Assert.False(textEffects.HasEffects);
            Assert.False(textEffects.HasShadow);
            Assert.False(textEffects.HasGlow);
            Assert.False(textEffects.HasReflection);
            Assert.False(textEffects.HasSoftEdge);
            Assert.Null(textEffects.Shadow);
            Assert.Null(textEffects.Glow);
            Assert.Null(textEffects.Reflection);
            Assert.Equal(0.0f, textEffects.SoftEdgeRadius);
        }

        [Fact]
        public void TestTextEffectsWithShadow()
        {
            // Create a TextEffectsInfo with shadow
            var textEffects = new TextEffectsInfo
            {
                HasEffects = true,
                HasShadow = true,
                Shadow = new ShadowInfo
                {
                    HasShadow = true,
                    Type = ShadowType.Outer,
                    Blur = 4.0f,
                    Distance = 3.0f,
                    Angle = 45.0f,
                    Color = new ColorInfo { Red = 0, Green = 0, Blue = 0 },
                    Transparency = 50.0f
                }
            };
            
            // Verify shadow properties
            Assert.True(textEffects.HasEffects);
            Assert.True(textEffects.HasShadow);
            Assert.NotNull(textEffects.Shadow);
            Assert.Equal(ShadowType.Outer, textEffects.Shadow.Type);
            Assert.Equal(4.0f, textEffects.Shadow.Blur);
            Assert.Equal(3.0f, textEffects.Shadow.Distance);
            Assert.Equal(45.0f, textEffects.Shadow.Angle);
        }

        [Fact]
        public void TestTextEffectsWithGlow()
        {
            // Create a TextEffectsInfo with glow
            var textEffects = new TextEffectsInfo
            {
                HasEffects = true,
                HasGlow = true,
                Glow = new GlowInfo
                {
                    Radius = 5.0f,
                    Color = new ColorInfo { Red = 255, Green = 255, Blue = 0 },
                    Transparency = 0.0f
                }
            };
            
            // Verify glow properties
            Assert.True(textEffects.HasEffects);
            Assert.True(textEffects.HasGlow);
            Assert.NotNull(textEffects.Glow);
            Assert.Equal(5.0f, textEffects.Glow.Radius);
            Assert.Equal(255, textEffects.Glow.Color.Red);
            Assert.Equal(255, textEffects.Glow.Color.Green);
            Assert.Equal(0, textEffects.Glow.Color.Blue);
        }

        [Fact]
        public void TestTextEffectsWithReflection()
        {
            // Create a TextEffectsInfo with reflection
            var textEffects = new TextEffectsInfo
            {
                HasEffects = true,
                HasReflection = true,
                Reflection = new ReflectionInfo
                {
                    BlurRadius = 2.0f,
                    StartOpacity = 1.0f,
                    EndAlpha = 0.0f,
                    Distance = 1.0f
                }
            };
            
            // Verify reflection properties
            Assert.True(textEffects.HasEffects);
            Assert.True(textEffects.HasReflection);
            Assert.NotNull(textEffects.Reflection);
            Assert.Equal(2.0f, textEffects.Reflection.BlurRadius);
            Assert.Equal(1.0f, textEffects.Reflection.StartOpacity);
            Assert.Equal(0.0f, textEffects.Reflection.EndAlpha);
        }

        [Fact]
        public void TestTextEffectsWithSoftEdge()
        {
            // Create a TextEffectsInfo with soft edge
            var textEffects = new TextEffectsInfo
            {
                HasEffects = true,
                HasSoftEdge = true,
                SoftEdgeRadius = 3.0f
            };
            
            // Verify soft edge properties
            Assert.True(textEffects.HasEffects);
            Assert.True(textEffects.HasSoftEdge);
            Assert.Equal(3.0f, textEffects.SoftEdgeRadius);
        }

        [Fact]
        public void TestTextEffectsWithMultipleEffects()
        {
            // Create a TextEffectsInfo with multiple effects
            var textEffects = new TextEffectsInfo
            {
                HasEffects = true,
                HasShadow = true,
                Shadow = new ShadowInfo
                {
                    HasShadow = true,
                    Type = ShadowType.Outer,
                    Blur = 4.0f
                },
                HasGlow = true,
                Glow = new GlowInfo
                {
                    Radius = 5.0f,
                    Color = new ColorInfo { Red = 255, Green = 255, Blue = 0 }
                }
            };
            
            // Verify multiple effects
            Assert.True(textEffects.HasEffects);
            Assert.True(textEffects.HasShadow);
            Assert.True(textEffects.HasGlow);
            Assert.NotNull(textEffects.Shadow);
            Assert.NotNull(textEffects.Glow);
        }

        [Fact]
        public void TestTextRunInfoWithTextEffects()
        {
            // Create a TextRunInfo with text effects
            var textRun = new TextRunInfo
            {
                Text = "Test Text",
                FontName = "Arial",
                FontSize = 12,
                TextEffects = new TextEffectsInfo
                {
                    HasEffects = true,
                    HasShadow = true,
                    Shadow = new ShadowInfo
                    {
                        HasShadow = true,
                        Type = ShadowType.Outer,
                        Blur = 4.0f,
                        Distance = 3.0f,
                        Angle = 45.0f
                    }
                }
            };
            
            // Verify text run has effects
            Assert.NotNull(textRun.TextEffects);
            Assert.True(textRun.TextEffects.HasEffects);
            Assert.True(textRun.TextEffects.HasShadow);
            Assert.NotNull(textRun.TextEffects.Shadow);
        }
    }
}
