using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Api.Excel;
using OfficeHelperOpenXml.Api.Word;
using OfficeHelperOpenXml.Elements;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Test
{
    class Program
    {
        static string TemplatesPath = "D:/pythonf/office_helper/OfficeHelper/examples/templates";

        static void Main(string[] args)
        {
            Console.WriteLine("╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║         OfficeHelperOpenXml - Phase 6: Full Testing          ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝");
            Console.WriteLine($"Version: {OfficeHelperWrapper.GetVersion()}\n");

            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "--ppt": TestAllPPT(); return;
                    case "--excel": TestAllExcel(); return;
                    case "--word": TestAllWord(); return;
                    case "--write": TestWriteFunctions(); return;
                    case "--perf": TestPerformance(); return;
                    case "--full": RunFullTest(); return;
                    case "--json": TestJsonOutput(); return;
                    case "--analyze": AnalyzeUserFiles(); return;
                    case "--conn": AnalyzeConnection(); return;
                    case "--slidewriter-json": SlideWriterJsonMethodsTest.RunTests(); return;
                    default:
                        if (File.Exists(args[0])) { TestSingleFile(args[0]); return; }
                        break;
                }
            }
            RunFullTest();
        }

        static void RunFullTest()
        {
            var sw = Stopwatch.StartNew();
            int passed = 0, failed = 0;

            Console.WriteLine("\n═══════════════════ 1. PPT READ TESTS ═══════════════════\n");
            var pptResults = TestAllPPT();
            passed += pptResults.Item1; failed += pptResults.Item2;

            Console.WriteLine("\n═══════════════════ 2. EXCEL READ TESTS ═══════════════════\n");
            var excelResults = TestAllExcel();
            passed += excelResults.Item1; failed += excelResults.Item2;

            Console.WriteLine("\n═══════════════════ 3. WORD READ TESTS ═══════════════════\n");
            var wordResults = TestAllWord();
            passed += wordResults.Item1; failed += wordResults.Item2;

            Console.WriteLine("\n═══════════════════ 4. WRITE FUNCTION TESTS ═══════════════════\n");
            var writeResults = TestWriteFunctions();
            passed += writeResults.Item1; failed += writeResults.Item2;

            Console.WriteLine("\n═══════════════════ 5. PERFORMANCE TESTS ═══════════════════\n");
            TestPerformance();

            sw.Stop();
            Console.WriteLine("\n╔══════════════════════════════════════════════════════════════╗");
            Console.WriteLine($"║  TOTAL: {passed} PASSED, {failed} FAILED   |   Time: {sw.ElapsedMilliseconds}ms");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════╝");
        }

        static (int, int) TestAllPPT()
        {
            int passed = 0, failed = 0;
            var pptFiles = new[] { "26xdemo1.pptx", "group.pptx", "picture.pptx", "table.pptx" };
            
            foreach (var file in pptFiles)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) { Console.WriteLine($"  [SKIP] {file} not found"); continue; }
                
                try
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new PowerPointReader())
                    {
                        reader.Load(path);
                        var json = reader.ToJson();
                        sw.Stop();
                        var info = reader.PresentationInfo;
                        Console.WriteLine($"  [PASS] {file}");
                        Console.WriteLine($"         Slides: {info.SlideCount}, Elements: {reader.GetAllElements().Count}, JSON: {json.Length} chars, Time: {sw.ElapsedMilliseconds}ms");
                        passed++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  [FAIL] {file}: {ex.Message}");
                    failed++;
                }
            }
            return (passed, failed);
        }

        static (int, int) TestAllExcel()
        {
            int passed = 0, failed = 0;
            var excelFiles = new[] { "26xdemo1.xlsx", "first.xlsx", "template_mozumingxi.xlsx" };
            
            foreach (var file in excelFiles)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) { Console.WriteLine($"  [SKIP] {file} not found"); continue; }
                
                try
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new ExcelReader())
                    {
                        reader.Load(path);
                        var sheets = reader.GetSheetNames();
                        var json = reader.ToJson();
                        sw.Stop();
                        Console.WriteLine($"  [PASS] {file}");
                        Console.WriteLine($"         Sheets: {sheets.Count} [{string.Join(", ", sheets.Take(3))}{(sheets.Count > 3 ? "..." : "")}], JSON: {json.Length} chars, Time: {sw.ElapsedMilliseconds}ms");
                        passed++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  [FAIL] {file}: {ex.Message}");
                    failed++;
                }
            }
            return (passed, failed);
        }

        static (int, int) TestAllWord()
        {
            int passed = 0, failed = 0;
            var wordFiles = new[] { "26xdemo1.docx" };
            
            foreach (var file in wordFiles)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) { Console.WriteLine($"  [SKIP] {file} not found"); continue; }
                
                try
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new WordReader())
                    {
                        reader.Load(path);
                        var paras = reader.GetParagraphCount();
                        var tables = reader.GetTableCount();
                        var json = reader.ToJson();
                        sw.Stop();
                        Console.WriteLine($"  [PASS] {file}");
                        Console.WriteLine($"         Paragraphs: {paras}, Tables: {tables}, JSON: {json.Length} chars, Time: {sw.ElapsedMilliseconds}ms");
                        passed++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  [FAIL] {file}: {ex.Message}");
                    failed++;
                }
            }
            return (passed, failed);
        }

        static (int, int) TestWriteFunctions()
        {
            int passed = 0, failed = 0;
            string tempDir = Path.GetTempPath();

            // Test 1: PPT Write
            Console.WriteLine("  Testing PPT Write...");
            try
            {
                string path = Path.Combine(tempDir, "test_write.pptx");
                if (File.Exists(path)) File.Delete(path);
                using (var writer = new PowerPointWriter())
                {
                    writer.OpenOrCreate(path);
                    writer.AddSlide();
                    var textBox = new TextBoxElement { Name = "TestBox" };
                    var pos = new PositionComponent { Bounds = new ShapeBounds { X = 2, Y = 2, Width = 10, Height = 3 } };
                    textBox.AddComponent(pos);
                    var text = new TextComponent { HasText = true };
                    text.Paragraphs.Add(new ParagraphInfo());
                    text.Paragraphs[0].Runs.Add(new TextRunInfo { Text = "Test Content", FontSize = 24 });
                    textBox.AddComponent(text);
                    writer.AddElement(1, textBox);
                    writer.Save();
                }
                // Verify
                using (var reader = new PowerPointReader())
                {
                    reader.Load(path);
                    if (reader.PresentationInfo.SlideCount >= 1) { Console.WriteLine("  [PASS] PPT Write"); passed++; }
                    else { Console.WriteLine("  [FAIL] PPT Write: verification failed"); failed++; }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  [FAIL] PPT Write: {ex.Message}"); failed++; }

            // Test 2: Excel Write
            Console.WriteLine("  Testing Excel Write...");
            try
            {
                string path = Path.Combine(tempDir, "test_write.xlsx");
                if (File.Exists(path)) File.Delete(path);
                using (var writer = new ExcelWriter())
                {
                    writer.OpenOrCreate(path);
                    var data = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { {"Col1", "A"}, {"Col2", "B"} },
                        new Dictionary<string, object> { {"Col1", "C"}, {"Col2", "D"} }
                    };
                    writer.WriteData("Sheet1", data);
                    writer.Save();
                }
                // Verify
                using (var reader = new ExcelReader())
                {
                    reader.Load(path);
                    var rows = reader.GetSheetData("Sheet1");
                    if (rows.Count == 2) { Console.WriteLine("  [PASS] Excel Write"); passed++; }
                    else { Console.WriteLine($"  [FAIL] Excel Write: expected 2 rows, got {rows.Count}"); failed++; }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  [FAIL] Excel Write: {ex.Message}"); failed++; }

            // Test 3: Word Write
            Console.WriteLine("  Testing Word Write...");
            try
            {
                string path = Path.Combine(tempDir, "test_write.docx");
                if (File.Exists(path)) File.Delete(path);
                using (var writer = new WordWriter())
                {
                    writer.OpenOrCreate(path);
                    writer.AddHeading("Title", 1);
                    writer.AddParagraph("Paragraph 1");
                    writer.AddParagraph("Paragraph 2", true, false, 14);
                    writer.AddTable(new List<List<string>> { new List<string> { "A", "B" }, new List<string> { "1", "2" } });
                    writer.Save();
                }
                // Verify
                using (var reader = new WordReader())
                {
                    reader.Load(path);
                    if (reader.GetParagraphCount() >= 3 && reader.GetTableCount() >= 1) { Console.WriteLine("  [PASS] Word Write"); passed++; }
                    else { Console.WriteLine("  [FAIL] Word Write: verification failed"); failed++; }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  [FAIL] Word Write: {ex.Message}"); failed++; }

            return (passed, failed);
        }

        static void TestPerformance()
        {
            Console.WriteLine("  Performance comparison (5 iterations each):\n");
            var files = new[]
            {
                ("26xdemo1.pptx", "PPT Small"),
                ("group.pptx", "PPT Group"),
                ("table.pptx", "PPT Table"),
            };

            foreach (var (file, desc) in files)
            {
                var path = Path.Combine(TemplatesPath, file);
                if (!File.Exists(path)) continue;

                var times = new List<long>();
                for (int i = 0; i < 5; i++)
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new PowerPointReader())
                    {
                        reader.Load(path);
                        reader.ToJson();
                    }
                    sw.Stop();
                    times.Add(sw.ElapsedMilliseconds);
                }
                Console.WriteLine($"  {desc,-15} : First={times[0]}ms, Avg(2-5)={times.Skip(1).Average():F1}ms, Min={times.Min()}ms");
            }

            // Excel performance
            var excelPath = Path.Combine(TemplatesPath, "template_mozumingxi.xlsx");
            if (File.Exists(excelPath))
            {
                var times = new List<long>();
                for (int i = 0; i < 5; i++)
                {
                    var sw = Stopwatch.StartNew();
                    using (var reader = new ExcelReader())
                    {
                        reader.Load(excelPath);
                        reader.ToJson();
                    }
                    sw.Stop();
                    times.Add(sw.ElapsedMilliseconds);
                }
                Console.WriteLine($"  {"Excel Large",-15} : First={times[0]}ms, Avg(2-5)={times.Skip(1).Average():F1}ms, Min={times.Min()}ms");
            }
        }

        static void TestSingleFile(string path)
        {
            Console.WriteLine($"Testing: {path}\n");
            var ext = Path.GetExtension(path).ToLower();
            var sw = Stopwatch.StartNew();

            if (ext == ".pptx")
            {
                using (var reader = new PowerPointReader())
                {
                    reader.Load(path);
                    var json = reader.ToJson();
                    sw.Stop();
                    Console.WriteLine($"Slides: {reader.PresentationInfo.SlideCount}");
                    Console.WriteLine($"Elements: {reader.GetAllElements().Count}");
                    Console.WriteLine($"JSON length: {json.Length}");
                    Console.WriteLine($"Time: {sw.ElapsedMilliseconds}ms");

                    // Save full JSON to file
                    var outputPath = Path.ChangeExtension(path, ".json");
                    File.WriteAllText(outputPath, json);
                    Console.WriteLine($"Full JSON saved to: {outputPath}");

                    Console.WriteLine($"\nJSON preview:\n{(json.Length > 2000 ? json.Substring(0, 2000) + "..." : json)}");
                }
            }
            else if (ext == ".xlsx")
            {
                using (var reader = new ExcelReader())
                {
                    reader.Load(path);
                    var json = reader.ToJson();
                    sw.Stop();
                    Console.WriteLine($"Sheets: {string.Join(", ", reader.GetSheetNames())}");
                    Console.WriteLine($"JSON length: {json.Length}");
                    Console.WriteLine($"Time: {sw.ElapsedMilliseconds}ms");
                    Console.WriteLine($"\nJSON preview:\n{(json.Length > 2000 ? json.Substring(0, 2000) + "..." : json)}");
                }
            }
            else if (ext == ".docx")
            {
                using (var reader = new WordReader())
                {
                    reader.Load(path);
                    var json = reader.ToJson();
                    sw.Stop();
                    Console.WriteLine($"Paragraphs: {reader.GetParagraphCount()}");
                    Console.WriteLine($"Tables: {reader.GetTableCount()}");
                    Console.WriteLine($"JSON length: {json.Length}");
                    Console.WriteLine($"Time: {sw.ElapsedMilliseconds}ms");
                    Console.WriteLine($"\nJSON preview:\n{(json.Length > 2000 ? json.Substring(0, 2000) + "..." : json)}");
                }
            }
        }

        static void TestJsonOutput()
        {
            Console.WriteLine("\n═══════════════════ JSON FORMAT TEST ═══════════════════\n");
            
            var testFiles = new[] { "picture.pptx", "26xdemo1.pptx", "group.pptx" };
            
            foreach (var file in testFiles)
            {
                string pptPath = Path.Combine(TemplatesPath, file);
                string outputPath = $"D:/pythonf/c_sharp_project/officeHelperOpenxml/test_output_{Path.GetFileNameWithoutExtension(file)}.json";
                
                if (!File.Exists(pptPath))
                {
                    Console.WriteLine($"File not found: {pptPath}");
                    continue;
                }
                
                Console.WriteLine($"\n--- Testing {file} ---");
                
                using (var reader = new PowerPointReader())
                {
                    reader.Load(pptPath);
                    var json = reader.ToJson();
                    
                    File.WriteAllText(outputPath, json);
                    
                    Console.WriteLine($"✓ JSON saved to: {outputPath}");
                    Console.WriteLine($"✓ JSON length: {json.Length} chars");
                    Console.WriteLine($"✓ Slides: {reader.PresentationInfo.SlideCount}");
                    Console.WriteLine($"✓ Elements: {reader.GetAllElements().Count}");
                    
                    Console.WriteLine($"\n--- JSON Preview (first 1500 chars) ---");
                    Console.WriteLine(json.Substring(0, Math.Min(1500, json.Length)));
                    if (json.Length > 1500) Console.WriteLine("...");
                }
            }
        }
        
        static void AnalyzeUserFiles()
        {
            var files = new[]
            {
                @"D:\pythonf\office_helper\OfficeHelper\examples\templates\方案報告(改造-單機-連線)-版本20251106更新.pptx",
                @"D:\pythonf\office_helper\OfficeHelper\examples\templates\方案報告(新增&再製-單機-連線)-版本20251106更新.pptx",
                @"D:\download\263 CNC7.2螺紋孔小徑檢測機改造253PL-DFM-V3.1 LYG--0406.pptx",
                @"D:\download\263Forging尺寸檢測機再制再制-DFM-V3-20250711.pptx"
            };
            
            AnalyzePPT.AnalyzeFiles(files);
        }
        
        static void AnalyzeConnection()
        {
            var file = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\方案報告(改造-單機-連線)-版本20251106更新.pptx";
            AnalyzeConnectionShape.Analyze(file);
        }
    }
}
