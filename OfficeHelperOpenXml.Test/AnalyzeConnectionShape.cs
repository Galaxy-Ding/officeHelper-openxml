using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Test
{
    public class AnalyzeConnectionShape
    {
        public static void Analyze(string filePath)
        {
            Console.WriteLine("\n╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║          ConnectionShape Structure Analysis                  ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝\n");

            using (var doc = PresentationDocument.Open(filePath, false))
            {
                var presentationPart = doc.PresentationPart;
                
                // 查找包含 ConnectionShape 的幻灯片
                P.ConnectionShape? conn = null;
                P.ShapeTree? shapeTree = null;
                int slideIndex = 0;
                
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    slideIndex++;
                    var tree = slidePart.Slide?.CommonSlideData?.ShapeTree;
                    if (tree == null) continue;
                    
                    var connShapes = tree.ChildElements.OfType<P.ConnectionShape>().ToList();
                    if (connShapes.Count > 0)
                    {
                        conn = connShapes[0];
                        shapeTree = tree;
                        Console.WriteLine($"Found {connShapes.Count} ConnectionShape elements in slide {slideIndex}\n");
                        break;
                    }
                }
                
                if (conn == null)
                {
                    Console.WriteLine("No ConnectionShape found in any slide.");
                    return;
                }

                Console.WriteLine("=== ConnectionShape Type Information ===");
                Console.WriteLine($"Full Type: {conn.GetType().FullName}");
                Console.WriteLine($"Base Type: {conn.GetType().BaseType?.Name}");
                Console.WriteLine();

                Console.WriteLine("=== Properties ===");
                foreach (var prop in conn.GetType().GetProperties())
                {
                    Console.WriteLine($"  - {prop.Name} : {prop.PropertyType.Name}");
                }
                Console.WriteLine();

                Console.WriteLine("=== Child Elements ===");
                foreach (var child in conn.ChildElements)
                {
                    Console.WriteLine($"  - {child.GetType().Name}");
                }
                Console.WriteLine();

                // NonVisualConnectionShapeProperties
                Console.WriteLine("=== NonVisualConnectionShapeProperties ===");
                Console.WriteLine($"  Has NonVisualConnectionShapeProperties: {conn.NonVisualConnectionShapeProperties != null}");
                Console.WriteLine();

                // ShapeProperties
                Console.WriteLine("=== ShapeProperties ===");
                Console.WriteLine($"  Has ShapeProperties: {conn.ShapeProperties != null}");
                Console.WriteLine();

                // ShapeStyle
                Console.WriteLine("=== ShapeStyle ===");
                if (conn.ShapeStyle != null)
                {
                    Console.WriteLine("  Has ShapeStyle: Yes");
                }
                else
                {
                    Console.WriteLine("  Has ShapeStyle: No");
                }
                Console.WriteLine();

                Console.WriteLine("=== XML Output (first 2000 chars) ===");
                var xml = conn.OuterXml;
                Console.WriteLine(xml.Substring(0, Math.Min(2000, xml.Length)));
                if (xml.Length > 2000) Console.WriteLine("...");
                Console.WriteLine();

                // 对比 Shape
                Console.WriteLine("\n=== Comparison: ConnectionShape vs Shape ===");
                var shapes = shapeTree.ChildElements.OfType<P.Shape>().ToList();
                if (shapes.Count > 0)
                {
                    var shape = shapes[0];
                    Console.WriteLine($"Shape has ShapeProperties: {shape.ShapeProperties != null}");
                    Console.WriteLine($"ConnectionShape has ShapeProperties: {conn.ShapeProperties != null}");
                    Console.WriteLine($"Shape base type: {shape.GetType().BaseType?.Name}");
                    Console.WriteLine($"ConnectionShape base type: {conn.GetType().BaseType?.Name}");
                }
            }
        }
    }
}
