using System;
using Xunit;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Components;
using System.Reflection;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for TextOutline comparison logic
    /// Tests the IsSameTextOutline method to ensure proper comparison of all properties
    /// </summary>
    public class TextOutlineComparisonTest
    {
        // Helper method to invoke the private IsSameTextOutline method
        private bool InvokeIsSameTextOutline(TextOutlineInfo outline1, TextOutlineInfo outline2)
        {
            var textComponent = new TextComponent();
            var method = typeof(TextComponent).GetMethod("IsSameTextOutline", 
                BindingFlags.NonPublic | BindingFlags.Instance);
            return (bool)method.Invoke(textComponent, new object[] { outline1, outline2 });
        }

        [Fact]
        public void IsSameTextOutline_BothNull_ReturnsTrue()
        {
            // Both null should be considered equal
            Assert.True(InvokeIsSameTextOutline(null, null));
        }

        [Fact]
        public void IsSameTextOutline_OneNull_ReturnsFalse()
        {
            var outline = new TextOutlineInfo { HasOutline = true };
            
            // One null should not be equal
            Assert.False(InvokeIsSameTextOutline(outline, null));
            Assert.False(InvokeIsSameTextOutline(null, outline));
        }

        [Fact]
        public void IsSameTextOutline_BothNoOutline_ReturnsTrue()
        {
            var outline1 = new TextOutlineInfo { HasOutline = false };
            var outline2 = new TextOutlineInfo { HasOutline = false };
            
            // Both with no outline should be equal
            Assert.True(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentHasOutline_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo { HasOutline = true };
            var outline2 = new TextOutlineInfo { HasOutline = false };
            
            // Different HasOutline values should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_IdenticalOutlines_ReturnsTrue()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                DashStyle = LineDashStyle.Dash,
                CompoundLineType = "Double",
                CapType = "Round",
                JoinType = "Miter",
                Transparency = 0.5f
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                DashStyle = LineDashStyle.Dash,
                CompoundLineType = "Double",
                CapType = "Round",
                JoinType = "Miter",
                Transparency = 0.5f
            };
            
            // Identical outlines should be equal
            Assert.True(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentWidth_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false)
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 3.0f,
                Color = new ColorInfo(255, 0, 0, false)
            };
            
            // Different width should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_SlightlyDifferentWidth_ReturnsTrue()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false)
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.005f, // Within epsilon (0.01f)
                Color = new ColorInfo(255, 0, 0, false)
            };
            
            // Width within epsilon should be equal
            Assert.True(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentColor_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false)
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(0, 255, 0, false)
            };
            
            // Different color should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentDashStyle_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                DashStyle = LineDashStyle.Solid
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                DashStyle = LineDashStyle.Dash
            };
            
            // Different dash style should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentCompoundLineType_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                CompoundLineType = "Single"
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                CompoundLineType = "Double"
            };
            
            // Different compound line type should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentCapType_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                CapType = "Flat"
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                CapType = "Round"
            };
            
            // Different cap type should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentJoinType_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                JoinType = "Round"
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                JoinType = "Miter"
            };
            
            // Different join type should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_DifferentTransparency_ReturnsFalse()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                Transparency = 0.0f
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                Transparency = 0.5f
            };
            
            // Different transparency should not be equal
            Assert.False(InvokeIsSameTextOutline(outline1, outline2));
        }

        [Fact]
        public void IsSameTextOutline_SlightlyDifferentTransparency_ReturnsTrue()
        {
            var outline1 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                Transparency = 0.5f
            };
            
            var outline2 = new TextOutlineInfo
            {
                HasOutline = true,
                Width = 2.0f,
                Color = new ColorInfo(255, 0, 0, false),
                Transparency = 0.505f // Within epsilon (0.01f)
            };
            
            // Transparency within epsilon should be equal
            Assert.True(InvokeIsSameTextOutline(outline1, outline2));
        }
    }
}
