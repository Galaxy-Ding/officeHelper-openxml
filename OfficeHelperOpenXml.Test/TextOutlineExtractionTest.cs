using System;
using Xunit;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Manual test to verify text outline extraction functionality
    /// </summary>
    public class TextOutlineExtractionTest
    {
        [Fact]
        public void TestTextOutlineInfo_DefaultValues()
        {
            // Create a default TextOutlineInfo
            var textOutline = new TextOutlineInfo();
            
            // Verify default values
            Assert.False(textOutline.HasOutline);
            Assert.Equal(0.0f, textOutline.Width);
            Assert.NotNull(textOutline.Color);
            Assert.Equal(LineDashStyle.Solid, textOutline.DashStyle);
            Assert.Equal("Single", textOutline.CompoundLineType);
            Assert.Equal("Flat", textOutline.CapType);
            Assert.Equal("Round", textOutline.JoinType);
            Assert.Equal(0.0f, textOutline.Transparency);
        }
        
        [Fact]
        public void TestTextOutlineInfo_WithOutline()
        {
            // Create an outline
            var textOutline = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(0, 0, 255, false),
                DashStyle = LineDashStyle.Dash,
                CompoundLineType = "Double",
                CapType = "Round",
                JoinType = "Miter",
                Transparency = 0.0f
            };
            
            // Verify outline properties
            Assert.True(textOutline.HasOutline);
            Assert.Equal(2.0f, textOutline.Width);
            Assert.Equal(0, textOutline.Color.Red);
            Assert.Equal(0, textOutline.Color.Green);
            Assert.Equal(255, textOutline.Color.Blue);
            Assert.Equal(LineDashStyle.Dash, textOutline.DashStyle);
            Assert.Equal("Double", textOutline.CompoundLineType);
            Assert.Equal("Round", textOutline.CapType);
            Assert.Equal("Miter", textOutline.JoinType);
            Assert.Equal(0.0f, textOutline.Transparency);
        }
        
        [Fact]
        public void TestTextRunInfo_HasTextOutlineProperty()
        {
            // Create a TextRunInfo
            var runInfo = new TextRunInfo();
            
            // Verify that TextOutline property exists and is initialized
            Assert.NotNull(runInfo.TextOutline);
            Assert.False(runInfo.TextOutline.HasOutline);
            Assert.Equal(0.0f, runInfo.TextOutline.Width);
        }
        
        [Fact]
        public void TestTextOutlineInfo_ToString()
        {
            // Test ToString for outline with properties
            var textOutline = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 1.5f,
                Color = new ColorInfo(255, 0, 0, false),
                DashStyle = LineDashStyle.Solid
            };
            
            string result = textOutline.ToString();
            Assert.Contains("Outline:", result);
            Assert.Contains("1.50pt", result);
            
            // Test ToString for no outline
            var noOutline = new TextOutlineInfo();
            Assert.Equal("No Outline", noOutline.ToString());
        }
    }
}
