using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Core.Readers;

namespace OfficeHelperOpenXml.Test
{
    public class AnalyzePPT
    {
        public static void AnalyzeFiles(string[] files)
        {
            Console.WriteLine("\n╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║              PPT Element Type Analysis                        ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝\n");

            var allElementTypes = new Dictionary<string, int>();
            var unknownElements = new HashSet<string>();

            foreach (var file in files)
            {
                if (!System.IO.File.Exists(file))
                {
                    Console.WriteLine($"✗ File not found: {file}\n");
                    continue;
                }

                Console.WriteLine($"--- Analyzing: {System.IO.Path.GetFileName(file)} ---");
                
                try
                {
                    using (var doc = PresentationDocument.Open(file, false))
                    {
                        var presentationPart = doc.PresentationPart;
                        if (presentationPart == null) continue;

                        int slideCount = 0;
                        var fileElementTypes = new Dictionary<string, int>();

                        // 分析普通幻灯片
                        foreach (var slidePart in presentationPart.SlideParts)
                        {
                            slideCount++;
                            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
                            if (shapeTree == null) continue;

                            AnalyzeShapeTree(shapeTree, fileElementTypes, unknownElements);
                        }

                        // 分析母版
                        foreach (var masterPart in presentationPart.SlideMasterParts)
                        {
                            var shapeTree = masterPart.SlideMaster?.CommonSlideData?.ShapeTree;
                            if (shapeTree != null)
                            {
                                AnalyzeShapeTree(shapeTree, fileElementTypes, unknownElements);
                            }

                            // 分析布局
                            foreach (var layoutPart in masterPart.SlideLayoutParts)
                            {
                                var layoutTree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree;
                                if (layoutTree != null)
                                {
                                    AnalyzeShapeTree(layoutTree, fileElementTypes, unknownElements);
                                }
                            }
                        }

                        Console.WriteLine($"  Slides: {slideCount}");
                        Console.WriteLine("  Element types found:");
                        foreach (var kvp in fileElementTypes.OrderByDescending(x => x.Value))
                        {
                            Console.WriteLine($"    - {kvp.Key}: {kvp.Value}");
                            if (!allElementTypes.ContainsKey(kvp.Key))
                                allElementTypes[kvp.Key] = 0;
                            allElementTypes[kvp.Key] += kvp.Value;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  ✗ Error: {ex.Message}");
                }

                Console.WriteLine();
            }

            // 总结
            Console.WriteLine("\n╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║                    Summary                                    ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝\n");

            Console.WriteLine("All element types across all files:");
            foreach (var kvp in allElementTypes.OrderByDescending(x => x.Value))
            {
                Console.WriteLine($"  - {kvp.Key}: {kvp.Value}");
            }

            if (unknownElements.Count > 0)
            {
                Console.WriteLine("\n⚠ Unknown/Unhandled element types:");
                foreach (var elem in unknownElements.OrderBy(x => x))
                {
                    Console.WriteLine($"  - {elem}");
                }
            }
            else
            {
                Console.WriteLine("\n✓ All element types are recognized!");
            }
        }

        private static void AnalyzeShapeTree(ShapeTree shapeTree, Dictionary<string, int> elementTypes, HashSet<string> unknownElements)
        {
            foreach (var child in shapeTree.ChildElements)
            {
                var typeName = child.GetType().Name;
                
                // 统计已知类型
                if (child is Shape)
                {
                    IncrementCount(elementTypes, "Shape");
                }
                else if (child is Picture)
                {
                    IncrementCount(elementTypes, "Picture");
                }
                else if (child is GroupShape groupShape)
                {
                    IncrementCount(elementTypes, "GroupShape");
                    // 递归分析组内元素
                    AnalyzeGroupShape(groupShape, elementTypes, unknownElements);
                }
                else if (child is GraphicFrame)
                {
                    IncrementCount(elementTypes, "GraphicFrame");
                }
                else if (child is ConnectionShape)
                {
                    IncrementCount(elementTypes, "ConnectionShape");
                    unknownElements.Add("ConnectionShape");
                }
                else if (child is NonVisualGroupShapeProperties || 
                         child is GroupShapeProperties ||
                         typeName == "ShapeTree")
                {
                    // 这些是结构性元素，不是内容元素
                    continue;
                }
                else
                {
                    IncrementCount(elementTypes, typeName);
                    unknownElements.Add(typeName);
                }
            }
        }

        private static void AnalyzeGroupShape(GroupShape groupShape, Dictionary<string, int> elementTypes, HashSet<string> unknownElements)
        {
            foreach (var child in groupShape.ChildElements)
            {
                var typeName = child.GetType().Name;
                
                if (child is Shape)
                {
                    IncrementCount(elementTypes, "Shape (in group)");
                }
                else if (child is Picture)
                {
                    IncrementCount(elementTypes, "Picture (in group)");
                }
                else if (child is GroupShape nestedGroup)
                {
                    IncrementCount(elementTypes, "GroupShape (nested)");
                    AnalyzeGroupShape(nestedGroup, elementTypes, unknownElements);
                }
                else if (child is GraphicFrame)
                {
                    IncrementCount(elementTypes, "GraphicFrame (in group)");
                }
                else if (child is ConnectionShape)
                {
                    IncrementCount(elementTypes, "ConnectionShape (in group)");
                    unknownElements.Add("ConnectionShape");
                }
                else if (child is NonVisualGroupShapeProperties || 
                         child is GroupShapeProperties)
                {
                    continue;
                }
                else
                {
                    IncrementCount(elementTypes, $"{typeName} (in group)");
                    unknownElements.Add(typeName);
                }
            }
        }

        private static void IncrementCount(Dictionary<string, int> dict, string key)
        {
            if (!dict.ContainsKey(key))
                dict[key] = 0;
            dict[key]++;
        }
    }
}
