using System;
using System.Collections.Generic;
using Xunit;
using FsCheck;
using FsCheck.Xunit;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Property-based tests for TextComponent JSON serialization
    /// Feature: remove-text-run-shadow
    /// </summary>
    public class TextComponentPropertyTests
    {
        /// <summary>
        /// Feature: remove-text-run-shadow, Property 1: Text run JSON excludes shadow
        /// Validates: Requirements 1.1
        /// 
        /// Property: For any text run serialized to JSON, the resulting JSON string 
        /// should not contain the substring "shadow" within the text run object.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property TextRunJsonExcludesShadow()
        {
            return Prop.ForAll(
                GenerateTextComponent(),
                textComponent =>
                {
                    // Skip null components (shouldn't happen but be defensive)
                    if (textComponent == null)
                        return true.ToProperty().Label("Null component skipped");
                    
                    // Serialize the text component to JSON
                    var json = textComponent.ToJson();
                    
                    // If there's no text, the property is trivially satisfied
                    if (!textComponent.HasText || textComponent.Paragraphs == null || textComponent.Paragraphs.Count == 0)
                        return true.ToProperty().Label("No text content");
                    
                    // Check that the text array doesn't contain shadow properties
                    // We need to verify that within the "text" array, no text run object contains "shadow"
                    
                    // Find the text array in the JSON
                    var textArrayStart = json.IndexOf("\"text\":[");
                    if (textArrayStart == -1)
                        return true.ToProperty().Label("No text array in JSON");
                    
                    // Extract everything after "text":[ until the end
                    // Since ToJson() returns just the component's JSON (not a complete object),
                    // we can simply check if the entire JSON contains "shadow" after the text array starts
                    var textArrayContent = json.Substring(textArrayStart + 8); // Skip past "text":[
                    
                    // Check that the text array content doesn't contain "shadow"
                    var hasShadow = textArrayContent.Contains("\"shadow\"");
                    
                    if (hasShadow)
                    {
                        // For debugging: show a snippet of the JSON
                        var snippet = textArrayContent.Length > 200 ? textArrayContent.Substring(0, 200) + "..." : textArrayContent;
                        return false.ToProperty().Label($"FAIL: Found shadow in text array. JSON snippet: {snippet}");
                    }
                    
                    return true.ToProperty().Label("PASS: No shadow in text array");
                });
        }
        
        /// <summary>
        /// Generator for TextComponent with random text runs
        /// </summary>
        private static Arbitrary<TextComponent> GenerateTextComponent()
        {
            // Generate simple non-empty strings for text content
            var simpleTextGen = Gen.Elements("Hello", "World", "Test", "Sample", "阴影", "文本", "测试");
            
            var textRunGen = from text in simpleTextGen
                            from fontName in Gen.Elements("Arial", "Calibri", "Times New Roman", "等线", "宋体")
                            from fontSize in Gen.Choose(8, 72)
                            from isBold in Arb.Default.Bool().Generator
                            from isItalic in Arb.Default.Bool().Generator
                            from isUnderline in Arb.Default.Bool().Generator
                            from isStrikethrough in Arb.Default.Bool().Generator
                            from r in Gen.Choose(0, 255)
                            from g in Gen.Choose(0, 255)
                            from b in Gen.Choose(0, 255)
                            select new TextRunInfo
                            {
                                Text = text,
                                FontName = fontName,
                                FontSize = fontSize,
                                IsBold = isBold,
                                IsItalic = isItalic,
                                IsUnderline = isUnderline,
                                IsStrikethrough = isStrikethrough,
                                FontColor = new ColorInfo { Red = r, Green = g, Blue = b },
                                HasShadow = false, // Text runs should not have shadows
                                Shadow = new ShadowInfo() // Initialize to empty shadow
                            };
            
            var paragraphGen = from runs in Gen.NonEmptyListOf(textRunGen)
                              select new ParagraphInfo
                              {
                                  Runs = new List<TextRunInfo>(runs),
                                  Alignment = TextAlignment.Left,
                                  Level = 0
                              };
            
            var textComponentGen = from paragraphs in Gen.NonEmptyListOf(paragraphGen)
                                  select CreateTextComponent(paragraphs);
            
            return Arb.From(textComponentGen);
        }
        
        /// <summary>
        /// Helper method to create a TextComponent from paragraphs
        /// </summary>
        private static TextComponent CreateTextComponent(IEnumerable<ParagraphInfo> paragraphs)
        {
            var component = new TextComponent
            {
                HasText = true,
                Paragraphs = new List<ParagraphInfo>(paragraphs),
                IsEnabled = true
            };
            
            // Set default font properties from first run if available
            var firstPara = component.Paragraphs.Count > 0 ? component.Paragraphs[0] : null;
            var firstRun = firstPara?.Runs.Count > 0 ? firstPara.Runs[0] : null;
            if (firstRun != null)
            {
                component.FontName = firstRun.FontName;
                component.FontSize = firstRun.FontSize;
                component.FontColor = firstRun.FontColor;
            }
            
            return component;
        }
        
        /// <summary>
        /// Feature: remove-text-run-shadow, Property 2: Shape-level shadow preservation
        /// Validates: Requirements 1.2
        /// 
        /// Property: For any shape with shadow effects, the shape-level JSON should contain 
        /// a "shadow" property with complete shadow information.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property ShapeLevelShadowPreservation()
        {
            return Prop.ForAll(
                GenerateShadowComponent(),
                shadowComponent =>
                {
                    // Skip null components
                    if (shadowComponent == null)
                        return true.ToProperty().Label("Null component skipped");
                    
                    // Only test components that have shadows
                    if (!shadowComponent.Shadow.HasShadow)
                        return true.ToProperty().Label("No shadow to preserve");
                    
                    // Serialize the shadow component to JSON
                    var json = shadowComponent.ToJson();
                    
                    // Verify that the JSON is not null
                    if (json == "null" || string.IsNullOrEmpty(json))
                        return false.ToProperty().Label("FAIL: Shadow component returned null JSON");
                    
                    // Verify that the JSON contains the required shadow properties
                    var hasHasShadow = json.Contains("\"has_shadow\"");
                    var hasColor = json.Contains("\"color\"");
                    var hasBlur = json.Contains("\"blur\"");
                    var hasShadowType = json.Contains("\"shadow_type\"");
                    
                    if (!hasHasShadow)
                        return false.ToProperty().Label("FAIL: Missing has_shadow property");
                    
                    if (!hasColor)
                        return false.ToProperty().Label("FAIL: Missing color property");
                    
                    if (!hasBlur)
                        return false.ToProperty().Label("FAIL: Missing blur property");
                    
                    if (!hasShadowType)
                        return false.ToProperty().Label("FAIL: Missing shadow_type property");
                    
                    // Verify has_shadow is set to 1 (true)
                    if (!json.Contains("\"has_shadow\":1"))
                        return false.ToProperty().Label("FAIL: has_shadow is not set to 1");
                    
                    return true.ToProperty().Label("PASS: Shape-level shadow preserved with all properties");
                });
        }
        
        /// <summary>
        /// Generator for ShadowComponent with random shadow properties
        /// </summary>
        private static Arbitrary<ShadowComponent> GenerateShadowComponent()
        {
            var shadowComponentGen = from hasShadow in Arb.Default.Bool().Generator
                                    from shadowType in Gen.Elements(ShadowType.Outer, ShadowType.Inner)
                                    from r in Gen.Choose(0, 255)
                                    from g in Gen.Choose(0, 255)
                                    from b in Gen.Choose(0, 255)
                                    from blur in Gen.Choose(0, 100).Select(x => (float)x)
                                    from distance in Gen.Choose(0, 100).Select(x => (float)x)
                                    from angle in Gen.Choose(0, 360).Select(x => (float)x)
                                    from transparency in Gen.Choose(0, 100).Select(x => (float)x)
                                    select CreateShadowComponent(
                                        hasShadow, 
                                        shadowType, 
                                        r, g, b, 
                                        blur, 
                                        distance, 
                                        angle, 
                                        transparency);
            
            return Arb.From(shadowComponentGen);
        }
        
        /// <summary>
        /// Helper method to create a ShadowComponent with specified properties
        /// </summary>
        private static ShadowComponent CreateShadowComponent(
            bool hasShadow,
            ShadowType shadowType,
            int r, int g, int b,
            float blur,
            float distance,
            float angle,
            float transparency)
        {
            var component = new ShadowComponent
            {
                IsEnabled = true,
                Shadow = new ShadowInfo
                {
                    HasShadow = hasShadow,
                    Type = shadowType,
                    ShadowTypeName = shadowType == ShadowType.Outer ? "outer" : "inner",
                    Color = new ColorInfo { Red = r, Green = g, Blue = b, IsTransparent = false },
                    Blur = blur,
                    Distance = distance,
                    Angle = angle,
                    Transparency = transparency,
                    Opacity = 100 - transparency,
                    Style = ShadowStyle.Custom
                }
            };
            
            return component;
        }
        
        /// <summary>
        /// Feature: remove-text-run-shadow, Property 3: Text formatting preservation
        /// Validates: Requirements 1.3
        /// 
        /// Property: For any text run with formatting properties (font, size, color, bold, italic, 
        /// underline, strikethrough), all these properties should be present in the serialized JSON output.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property TextFormattingPreservation()
        {
            return Prop.ForAll(
                GenerateTextComponent(),
                textComponent =>
                {
                    // Skip null components
                    if (textComponent == null)
                        return true.ToProperty().Label("Null component skipped");
                    
                    // Skip components without text
                    if (!textComponent.HasText || textComponent.Paragraphs == null || textComponent.Paragraphs.Count == 0)
                        return true.ToProperty().Label("No text content");
                    
                    // Serialize the text component to JSON
                    var json = textComponent.ToJson();
                    
                    // Verify that all required formatting properties are present in the JSON
                    var requiredProperties = new[]
                    {
                        "\"content\":",
                        "\"font\":",
                        "\"font_size\":",
                        "\"font_color\":",
                        "\"font_bold\":",
                        "\"font_italic\":",
                        "\"font_underline\":",
                        "\"font_strikethrough\":"
                    };
                    
                    foreach (var property in requiredProperties)
                    {
                        if (!json.Contains(property))
                        {
                            return false.ToProperty().Label($"FAIL: Missing required property {property}");
                        }
                    }
                    
                    // Verify that the formatting values match the original text runs
                    // Extract the first run from the component for validation
                    var firstRun = textComponent.Paragraphs[0].Runs[0];
                    
                    // Check that the JSON contains the font name
                    if (!json.Contains($"\"{firstRun.FontName}\""))
                    {
                        return false.ToProperty().Label($"FAIL: Font name {firstRun.FontName} not found in JSON");
                    }
                    
                    // Check that the JSON contains the font size
                    if (!json.Contains($"\"font_size\":{firstRun.FontSize:F1}"))
                    {
                        return false.ToProperty().Label($"FAIL: Font size {firstRun.FontSize} not found in JSON");
                    }
                    
                    // Check that the JSON contains the bold value
                    var boldValue = firstRun.IsBold ? 1 : 0;
                    if (!json.Contains($"\"font_bold\":{boldValue}"))
                    {
                        return false.ToProperty().Label($"FAIL: Bold value {boldValue} not found in JSON");
                    }
                    
                    // Check that the JSON contains the italic value
                    var italicValue = firstRun.IsItalic ? 1 : 0;
                    if (!json.Contains($"\"font_italic\":{italicValue}"))
                    {
                        return false.ToProperty().Label($"FAIL: Italic value {italicValue} not found in JSON");
                    }
                    
                    // Check that the JSON contains the underline value
                    var underlineValue = firstRun.IsUnderline ? 1 : 0;
                    if (!json.Contains($"\"font_underline\":{underlineValue}"))
                    {
                        return false.ToProperty().Label($"FAIL: Underline value {underlineValue} not found in JSON");
                    }
                    
                    // Check that the JSON contains the strikethrough value
                    var strikethroughValue = firstRun.IsStrikethrough ? 1 : 0;
                    if (!json.Contains($"\"font_strikethrough\":{strikethroughValue}"))
                    {
                        return false.ToProperty().Label($"FAIL: Strikethrough value {strikethroughValue} not found in JSON");
                    }
                    
                    return true.ToProperty().Label("PASS: All text formatting properties preserved");
                });
        }
        
        /// <summary>
        /// Feature: remove-text-run-shadow, Property 4: Consistent structure across text runs
        /// Validates: Requirements 1.4
        /// 
        /// Property: For any set of text runs within a shape, all text runs should have 
        /// the same JSON structure without shadow properties.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property ConsistentStructureAcrossTextRuns()
        {
            return Prop.ForAll(
                GenerateTextComponent(),
                textComponent =>
                {
                    // Skip null components
                    if (textComponent == null)
                        return true.ToProperty().Label("Null component skipped");
                    
                    // Skip components without text
                    if (!textComponent.HasText || textComponent.Paragraphs == null || textComponent.Paragraphs.Count == 0)
                        return true.ToProperty().Label("No text content");
                    
                    // Serialize the text component to JSON
                    var json = textComponent.ToJson();
                    
                    // Find the text array in the JSON
                    var textArrayStart = json.IndexOf("\"text\":[");
                    if (textArrayStart == -1)
                        return true.ToProperty().Label("No text array in JSON");
                    
                    // Extract the text array content
                    var textArrayContent = json.Substring(textArrayStart + 8); // Skip past "text":[
                    var textArrayEnd = textArrayContent.LastIndexOf(']');
                    if (textArrayEnd != -1)
                        textArrayContent = textArrayContent.Substring(0, textArrayEnd);
                    
                    // Split the text array into individual text run objects
                    // We'll use a simple approach: split by "},{" to get individual runs
                    var runs = textArrayContent.Split(new[] { "},{" }, StringSplitOptions.RemoveEmptyEntries);
                    
                    if (runs.Length == 0)
                        return true.ToProperty().Label("No text runs found");
                    
                    // Check that all runs have the same structure (same set of properties)
                    // We'll verify that all runs have the required properties and no shadow property
                    var requiredProperties = new[]
                    {
                        "\"content\":",
                        "\"font\":",
                        "\"font_size\":",
                        "\"font_color\":",
                        "\"font_bold\":",
                        "\"font_italic\":",
                        "\"font_underline\":",
                        "\"font_strikethrough\":"
                    };
                    
                    foreach (var run in runs)
                    {
                        // Check that all required properties are present
                        foreach (var property in requiredProperties)
                        {
                            if (!run.Contains(property))
                            {
                                return false.ToProperty().Label($"FAIL: Text run missing property {property}");
                            }
                        }
                        
                        // Check that shadow property is NOT present
                        if (run.Contains("\"shadow\":"))
                        {
                            return false.ToProperty().Label("FAIL: Text run contains shadow property");
                        }
                    }
                    
                    return true.ToProperty().Label("PASS: All text runs have consistent structure without shadow");
                });
        }
        
        /// <summary>
        /// Feature: remove-text-run-shadow, Property 5: Model-JSON alignment
        /// Validates: Requirements 2.1, 2.2, 2.3
        /// 
        /// Property: For any TextJsonData instance serialized to JSON, the JSON output should 
        /// only contain properties that exist in the TextJsonData class definition.
        /// </summary>
        [Property(MaxTest = 100)]
        public Property ModelJsonAlignment()
        {
            return Prop.ForAll(
                GenerateTextJsonData(),
                textJsonData =>
                {
                    // Skip null data
                    if (textJsonData == null)
                        return true.ToProperty().Label("Null data skipped");
                    
                    // Serialize the TextJsonData to JSON using Newtonsoft.Json
                    var json = Newtonsoft.Json.JsonConvert.SerializeObject(textJsonData);
                    
                    // Verify that the JSON only contains properties defined in TextJsonData
                    var allowedProperties = new[]
                    {
                        "\"content\":",
                        "\"font\":",
                        "\"font_size\":",
                        "\"font_color\":",
                        "\"font_bold\":",
                        "\"font_italic\":",
                        "\"font_underline\":",
                        "\"font_strikethrough\":"
                    };
                    
                    // Check that shadow property is NOT present
                    if (json.Contains("\"shadow\":") || json.Contains("\"Shadow\":"))
                    {
                        return false.ToProperty().Label("FAIL: JSON contains shadow property not in TextJsonData model");
                    }
                    
                    // Verify that all properties in the JSON are in the allowed list
                    // Extract all property names from the JSON
                    var propertyPattern = new System.Text.RegularExpressions.Regex("\"([^\"]+)\":");
                    var matches = propertyPattern.Matches(json);
                    
                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        var propertyName = match.Groups[1].Value;
                        var propertyWithColon = $"\"{propertyName}\":";
                        
                        // Check if this property is in the allowed list
                        var isAllowed = false;
                        foreach (var allowed in allowedProperties)
                        {
                            if (allowed == propertyWithColon)
                            {
                                isAllowed = true;
                                break;
                            }
                        }
                        
                        if (!isAllowed)
                        {
                            return false.ToProperty().Label($"FAIL: JSON contains unexpected property {propertyName}");
                        }
                    }
                    
                    return true.ToProperty().Label("PASS: JSON output aligns with TextJsonData model");
                });
        }
        
        /// <summary>
        /// Generator for TextJsonData with random properties
        /// </summary>
        private static Arbitrary<Models.Json.TextJsonData> GenerateTextJsonData()
        {
            var textJsonDataGen = from content in Gen.Elements("Hello", "World", "Test", "Sample", "阴影", "文本", "测试")
                                 from fontName in Gen.Elements("Arial", "Calibri", "Times New Roman", "等线", "宋体")
                                 from fontSize in Gen.Choose(8, 72).Select(x => (float)x)
                                 from isBold in Arb.Default.Bool().Generator
                                 from isItalic in Arb.Default.Bool().Generator
                                 from isUnderline in Arb.Default.Bool().Generator
                                 from isStrikethrough in Arb.Default.Bool().Generator
                                 from r in Gen.Choose(0, 255)
                                 from g in Gen.Choose(0, 255)
                                 from b in Gen.Choose(0, 255)
                                 select new Models.Json.TextJsonData
                                 {
                                     Content = content,
                                     Font = fontName,
                                     FontSize = fontSize,
                                     FontColor = $"RGB({r}, {g}, {b})",
                                     FontBold = isBold ? 1 : 0,
                                     FontItalic = isItalic ? 1 : 0,
                                     FontUnderline = isUnderline ? 1 : 0,
                                     FontStrikethrough = isStrikethrough ? 1 : 0
                                 };
            
            return Arb.From(textJsonDataGen);
        }
    }
}
