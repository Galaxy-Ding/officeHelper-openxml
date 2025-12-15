using System;
using Xunit;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Manual test to verify text fill extraction functionality
    /// </summary>
    public class TextFillExtractionTest
    {
        [Fact]
        public void TestTextFillInfo_DefaultValues()
        {
            // Create a default TextFillInfo
            var textFill = new TextFillInfo();
            
            // Verify default values
            Assert.False(textFill.HasFill);
            Assert.Equal(FillType.NoFill, textFill.FillType);
            Assert.NotNull(textFill.Color);
            Assert.Equal(0.0f, textFill.Transparency);
            Assert.Null(textFill.Gradient);
            Assert.Null(textFill.Pattern);
        }
        
        [Fact]
        public void TestTextFillInfo_SolidFill()
        {
            // Create a solid fill
            var textFill = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Solid,
                Color = new ColorInfo(255, 0, 0, false),
                Transparency = 0.0f
            };
            
            // Verify solid fill properties
            Assert.True(textFill.HasFill);
            Assert.Equal(FillType.Solid, textFill.FillType);
            Assert.Equal(255, textFill.Color.Red);
            Assert.Equal(0, textFill.Color.Green);
            Assert.Equal(0, textFill.Color.Blue);
            Assert.Equal(0.0f, textFill.Transparency);
        }
        
        [Fact]
        public void TestTextFillInfo_GradientFill()
        {
            // Create a gradient fill
            var gradientInfo = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f
            };
            gradientInfo.Stops.Add(new GradientStop(0.0f, new ColorInfo(255, 0, 0, false)));
            gradientInfo.Stops.Add(new GradientStop(1.0f, new ColorInfo(0, 0, 255, false)));
            
            var textFill = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Gradient,
                Gradient = gradientInfo
            };
            
            // Verify gradient fill properties
            Assert.True(textFill.HasFill);
            Assert.Equal(FillType.Gradient, textFill.FillType);
            Assert.NotNull(textFill.Gradient);
            Assert.Equal("Linear", textFill.Gradient.GradientType);
            Assert.Equal(90.0f, textFill.Gradient.Angle);
            Assert.Equal(2, textFill.Gradient.Stops.Count);
            Assert.Equal(0.0f, textFill.Gradient.Stops[0].Position);
            Assert.Equal(1.0f, textFill.Gradient.Stops[1].Position);
        }
        
        [Fact]
        public void TestTextFillInfo_PatternFill()
        {
            // Create a pattern fill
            var patternInfo = new PatternInfo
            {
                PatternType = "Dots",
                ForegroundColor = new ColorInfo(255, 0, 0, false),
                BackgroundColor = new ColorInfo(255, 255, 255, false)
            };
            
            var textFill = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Pattern,
                Pattern = patternInfo
            };
            
            // Verify pattern fill properties
            Assert.True(textFill.HasFill);
            Assert.Equal(FillType.Pattern, textFill.FillType);
            Assert.NotNull(textFill.Pattern);
            Assert.Equal("Dots", textFill.Pattern.PatternType);
            Assert.Equal(255, textFill.Pattern.ForegroundColor.Red);
            Assert.Equal(255, textFill.Pattern.BackgroundColor.Red);
        }
        
        [Fact]
        public void TestTextRunInfo_HasTextFillProperty()
        {
            // Create a TextRunInfo
            var runInfo = new TextRunInfo();
            
            // Verify that TextFill property exists and is initialized
            Assert.NotNull(runInfo.TextFill);
            Assert.False(runInfo.TextFill.HasFill);
            Assert.Equal(FillType.NoFill, runInfo.TextFill.FillType);
        }
    }
}
