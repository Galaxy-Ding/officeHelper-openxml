using Xunit;
using OfficeHelperOpenXml.Utils;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Test
{
    public class ColorHelperJsonTests
    {
        [Fact]
        public void ParseRgbToHex_ValidRgbString_ReturnsHex()
        {
            // Arrange
            string rgbString = "RGB(255, 128, 64)";

            // Act
            string result = ColorHelper.ParseRgbToHex(rgbString);

            // Assert
            Assert.Equal("FF8040", result);
        }

        [Fact]
        public void ParseRgbToHex_WithWhitespace_ReturnsHex()
        {
            // Arrange
            string rgbString = "  RGB( 100 , 200 , 50 )  ";

            // Act
            string result = ColorHelper.ParseRgbToHex(rgbString);

            // Assert
            Assert.Equal("64C832", result);
        }

        [Fact]
        public void ParseRgbToHex_LowercaseRgb_ReturnsHex()
        {
            // Arrange
            string rgbString = "rgb(10, 20, 30)";

            // Act
            string result = ColorHelper.ParseRgbToHex(rgbString);

            // Assert
            Assert.Equal("0A141E", result);
        }

        [Fact]
        public void ParseRgbToHex_InvalidFormat_ReturnsEmpty()
        {
            // Arrange
            string rgbString = "255, 128, 64";

            // Act
            string result = ColorHelper.ParseRgbToHex(rgbString);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void ParseRgbToHex_OutOfRange_ReturnsEmpty()
        {
            // Arrange
            string rgbString = "RGB(300, 128, 64)";

            // Act
            string result = ColorHelper.ParseRgbToHex(rgbString);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void ParseRgbToHex_NegativeValue_ReturnsEmpty()
        {
            // Arrange
            string rgbString = "RGB(-10, 128, 64)";

            // Act
            string result = ColorHelper.ParseRgbToHex(rgbString);

            // Assert
            Assert.Equal("", result);
        }

        [Fact]
        public void ParseRgbToHex_NullOrEmpty_ReturnsEmpty()
        {
            // Act & Assert
            Assert.Equal("", ColorHelper.ParseRgbToHex(null));
            Assert.Equal("", ColorHelper.ParseRgbToHex(""));
            Assert.Equal("", ColorHelper.ParseRgbToHex("   "));
        }

        [Fact]
        public void MapSchemeColorName_Bg1_ReturnsBackground1()
        {
            // Act
            var result = ColorHelper.MapSchemeColorName("bg1");

            // Assert
            Assert.Equal(A.SchemeColorValues.Background1, result);
        }

        [Fact]
        public void MapSchemeColorName_Tx1_ReturnsText1()
        {
            // Act
            var result = ColorHelper.MapSchemeColorName("tx1");

            // Assert
            Assert.Equal(A.SchemeColorValues.Text1, result);
        }

        [Fact]
        public void MapSchemeColorName_Accent1_ReturnsAccent1()
        {
            // Act
            var result = ColorHelper.MapSchemeColorName("accent1");

            // Assert
            Assert.Equal(A.SchemeColorValues.Accent1, result);
        }

        [Fact]
        public void MapSchemeColorName_AllAccents_ReturnsCorrectValues()
        {
            // Test all accent colors
            Assert.Equal(A.SchemeColorValues.Accent1, ColorHelper.MapSchemeColorName("accent1"));
            Assert.Equal(A.SchemeColorValues.Accent2, ColorHelper.MapSchemeColorName("accent2"));
            Assert.Equal(A.SchemeColorValues.Accent3, ColorHelper.MapSchemeColorName("accent3"));
            Assert.Equal(A.SchemeColorValues.Accent4, ColorHelper.MapSchemeColorName("accent4"));
            Assert.Equal(A.SchemeColorValues.Accent5, ColorHelper.MapSchemeColorName("accent5"));
            Assert.Equal(A.SchemeColorValues.Accent6, ColorHelper.MapSchemeColorName("accent6"));
        }

        [Fact]
        public void CreateSchemeColorWithTransforms_WithLumMod_AppliesTransform()
        {
            // Act
            var result = ColorHelper.CreateSchemeColorWithTransforms("accent1", 60000, null);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(A.SchemeColorValues.Accent1, result.Val.Value);
            var lumMod = result.GetFirstChild<A.LuminanceModulation>();
            Assert.NotNull(lumMod);
            Assert.Equal(60000, lumMod.Val.Value);
        }

        [Fact]
        public void CreateSchemeColorWithTransforms_WithLumOff_AppliesTransform()
        {
            // Act
            var result = ColorHelper.CreateSchemeColorWithTransforms("bg1", null, 40000);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(A.SchemeColorValues.Background1, result.Val.Value);
            var lumOff = result.GetFirstChild<A.LuminanceOffset>();
            Assert.NotNull(lumOff);
            Assert.Equal(40000, lumOff.Val.Value);
        }

        [Fact]
        public void CreateSchemeColorWithTransforms_WithBothTransforms_AppliesBoth()
        {
            // Act
            var result = ColorHelper.CreateSchemeColorWithTransforms("accent2", 60000, 40000);

            // Assert
            Assert.NotNull(result);
            Assert.Equal(A.SchemeColorValues.Accent2, result.Val.Value);
            
            var lumMod = result.GetFirstChild<A.LuminanceModulation>();
            Assert.NotNull(lumMod);
            Assert.Equal(60000, lumMod.Val.Value);
            
            var lumOff = result.GetFirstChild<A.LuminanceOffset>();
            Assert.NotNull(lumOff);
            Assert.Equal(40000, lumOff.Val.Value);
        }

        [Fact]
        public void CreateColorFromJson_WithSchemeColor_ReturnsSchemeColor()
        {
            // Act
            var result = ColorHelper.CreateColorFromJson("RGB(255, 0, 0)", "accent1", 60000, 40000);

            // Assert
            Assert.NotNull(result);
            Assert.IsType<A.SchemeColor>(result);
            var schemeColor = (A.SchemeColor)result;
            Assert.Equal(A.SchemeColorValues.Accent1, schemeColor.Val.Value);
        }

        [Fact]
        public void CreateColorFromJson_WithOnlyRgb_ReturnsRgbColor()
        {
            // Act
            var result = ColorHelper.CreateColorFromJson("RGB(255, 128, 64)", null, null, null);

            // Assert
            Assert.NotNull(result);
            Assert.IsType<A.RgbColorModelHex>(result);
            var rgbColor = (A.RgbColorModelHex)result;
            Assert.Equal("FF8040", rgbColor.Val.Value);
        }

        [Fact]
        public void CreateColorFromJson_SchemeColorPriority_IgnoresRgb()
        {
            // Act - Both RGB and schemeColor provided, schemeColor should take priority
            var result = ColorHelper.CreateColorFromJson("RGB(255, 0, 0)", "bg1", null, null);

            // Assert
            Assert.NotNull(result);
            Assert.IsType<A.SchemeColor>(result);
            var schemeColor = (A.SchemeColor)result;
            Assert.Equal(A.SchemeColorValues.Background1, schemeColor.Val.Value);
        }

        [Fact]
        public void CreateColorFromJson_NoValidColor_ReturnsNull()
        {
            // Act
            var result = ColorHelper.CreateColorFromJson(null, null, null, null);

            // Assert
            Assert.Null(result);
        }

        [Fact]
        public void CreateColorFromJson_InvalidRgb_ReturnsNull()
        {
            // Act
            var result = ColorHelper.CreateColorFromJson("invalid", null, null, null);

            // Assert
            Assert.Null(result);
        }
    }
}
