using Xunit;
using OfficeHelperOpenXml.Models;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests for text fill comparison logic (IsSameTextFill method)
    /// Validates Requirements 2.1: Deep equality for gradient and pattern fills
    /// </summary>
    public class TextFillComparisonTest
    {
        [Fact]
        public void TestSolidFillComparison_Identical()
        {
            // Arrange
            var fill1 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Solid,
                Color = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                Transparency = 0.0f
            };

            var fill2 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Solid,
                Color = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                Transparency = 0.0f
            };

            // Act & Assert - using reflection to access private method
            var textComponent = new Components.TextComponent();
            var method = typeof(Components.TextComponent).GetMethod("IsSameTextFill", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var result = (bool)method.Invoke(textComponent, new object[] { fill1, fill2 });

            Assert.True(result, "Identical solid fills should be equal");
        }

        [Fact]
        public void TestGradientFillComparison_Identical()
        {
            // Arrange
            var gradient1 = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f,
                Stops = new List<GradientStop>
                {
                    new GradientStop(0.0f, new ColorInfo { Red = 255, Green = 0, Blue = 0 }),
                    new GradientStop(1.0f, new ColorInfo { Red = 0, Green = 0, Blue = 255 })
                }
            };

            var gradient2 = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f,
                Stops = new List<GradientStop>
                {
                    new GradientStop(0.0f, new ColorInfo { Red = 255, Green = 0, Blue = 0 }),
                    new GradientStop(1.0f, new ColorInfo { Red = 0, Green = 0, Blue = 255 })
                }
            };

            var fill1 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Gradient,
                Gradient = gradient1,
                Transparency = 0.0f
            };

            var fill2 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Gradient,
                Gradient = gradient2,
                Transparency = 0.0f
            };

            // Act & Assert
            var textComponent = new Components.TextComponent();
            var method = typeof(Components.TextComponent).GetMethod("IsSameTextFill", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var result = (bool)method.Invoke(textComponent, new object[] { fill1, fill2 });

            Assert.True(result, "Identical gradient fills should be equal");
        }

        [Fact]
        public void TestGradientFillComparison_DifferentStops()
        {
            // Arrange
            var gradient1 = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f,
                Stops = new List<GradientStop>
                {
                    new GradientStop(0.0f, new ColorInfo { Red = 255, Green = 0, Blue = 0 }),
                    new GradientStop(1.0f, new ColorInfo { Red = 0, Green = 0, Blue = 255 })
                }
            };

            var gradient2 = new GradientInfo
            {
                GradientType = "Linear",
                Angle = 90.0f,
                Stops = new List<GradientStop>
                {
                    new GradientStop(0.0f, new ColorInfo { Red = 255, Green = 0, Blue = 0 }),
                    new GradientStop(1.0f, new ColorInfo { Red = 0, Green = 255, Blue = 0 }) // Different color
                }
            };

            var fill1 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Gradient,
                Gradient = gradient1,
                Transparency = 0.0f
            };

            var fill2 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Gradient,
                Gradient = gradient2,
                Transparency = 0.0f
            };

            // Act & Assert
            var textComponent = new Components.TextComponent();
            var method = typeof(Components.TextComponent).GetMethod("IsSameTextFill", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var result = (bool)method.Invoke(textComponent, new object[] { fill1, fill2 });

            Assert.False(result, "Gradient fills with different stop colors should not be equal");
        }

        [Fact]
        public void TestPatternFillComparison_Identical()
        {
            // Arrange
            var pattern1 = new PatternInfo
            {
                PatternType = "Dots",
                ForegroundColor = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                BackgroundColor = new ColorInfo { Red = 255, Green = 255, Blue = 255 }
            };

            var pattern2 = new PatternInfo
            {
                PatternType = "Dots",
                ForegroundColor = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                BackgroundColor = new ColorInfo { Red = 255, Green = 255, Blue = 255 }
            };

            var fill1 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Pattern,
                Pattern = pattern1,
                Transparency = 0.0f
            };

            var fill2 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Pattern,
                Pattern = pattern2,
                Transparency = 0.0f
            };

            // Act & Assert
            var textComponent = new Components.TextComponent();
            var method = typeof(Components.TextComponent).GetMethod("IsSameTextFill", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var result = (bool)method.Invoke(textComponent, new object[] { fill1, fill2 });

            Assert.True(result, "Identical pattern fills should be equal");
        }

        [Fact]
        public void TestPatternFillComparison_DifferentType()
        {
            // Arrange
            var pattern1 = new PatternInfo
            {
                PatternType = "Dots",
                ForegroundColor = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                BackgroundColor = new ColorInfo { Red = 255, Green = 255, Blue = 255 }
            };

            var pattern2 = new PatternInfo
            {
                PatternType = "Stripes", // Different pattern type
                ForegroundColor = new ColorInfo { Red = 255, Green = 0, Blue = 0 },
                BackgroundColor = new ColorInfo { Red = 255, Green = 255, Blue = 255 }
            };

            var fill1 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Pattern,
                Pattern = pattern1,
                Transparency = 0.0f
            };

            var fill2 = new TextFillInfo
            {
                HasFill = true,
                FillType = FillType.Pattern,
                Pattern = pattern2,
                Transparency = 0.0f
            };

            // Act & Assert
            var textComponent = new Components.TextComponent();
            var method = typeof(Components.TextComponent).GetMethod("IsSameTextFill", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var result = (bool)method.Invoke(textComponent, new object[] { fill1, fill2 });

            Assert.False(result, "Pattern fills with different types should not be equal");
        }

        [Fact]
        public void TestNullFillComparison()
        {
            // Act & Assert
            var textComponent = new Components.TextComponent();
            var method = typeof(Components.TextComponent).GetMethod("IsSameTextFill", 
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            
            var result1 = (bool)method.Invoke(textComponent, new object[] { null, null });
            Assert.True(result1, "Two null fills should be equal");

            var fill = new TextFillInfo { HasFill = false, FillType = FillType.NoFill };
            var result2 = (bool)method.Invoke(textComponent, new object[] { fill, null });
            Assert.False(result2, "Null and non-null fills should not be equal");
        }
    }
}
