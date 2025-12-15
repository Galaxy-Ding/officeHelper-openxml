using System;
using System.IO;
using System.Collections.Generic;
using OfficeHelperOpenXml.Api;
using OfficeHelperOpenXml.Api.Excel;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Core.Comparison;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeHelperOpenXml
{
    /// <summary>
    /// OfficeHelperOpenXml 主程序入口
    /// 用于直接调试 PowerPoint/Excel 分析功能
    /// 参考 D:\pythonf\office_helper\OfficeHelper\Program.cs 的调用方式
    /// </summary>
    class Program
    {
        /// <summary>
        /// 主入口点
        /// </summary>
        /// <param name="args">命令行参数</param>
        /// <returns>退出代码</returns>
        static int Main(string[] args)
        {
            try
            {
                // Check for command-line arguments
                if (args.Length > 0)
                {
                    return ParseCommandLineArguments(args);
                }

                // 显示欢迎信息
                Console.WriteLine("========================================");
                Console.WriteLine("  OfficeHelperOpenXml - 调试工具");
                Console.WriteLine("  基于 OpenXML SDK 的 Office 文件分析");
                Console.WriteLine("========================================");
                Console.WriteLine();

                // ============================================
                // 测试区域 - 根据需要取消注释相应的测试代码
                // ============================================

                // 测试 1: PowerPoint 文件分析
                //string pptPath = @"D:\pythonf\c_sharp_project\officeHelperOpenxml\test_ppt\textbox_2_master.pptx";
                //string outputJsonPath = @"D:\pythonf\c_sharp_project\officeHelperOpenxml\test_ppt\textbox_2_master.json";
                //return ProcessPowerPoint(pptPath, outputJsonPath);

                // 测试 2: Excel 文件分析
                //string excelPath = @"D:\test\sample.xlsx";
                //string excelOutputPath = @"D:\test\output_excel.json";
                //return ProcessExcel(excelPath, excelOutputPath);

                // 测试 3: 从 JSON 恢复 PowerPoint (旧方法)
                //string jsonPath = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\output_openxml.json";
                //string outputPptPath = @"D:\pythonf\office_helper\OfficeHelper\examples\templates\restored_openxml.pptx";
                //return RestorePPTFromJson(jsonPath, outputPptPath);

                 //============================================
                 //JSON to PPTX Conversion (Recommended)
                 //============================================
                 //This feature converts a JSON file containing presentation data
                 //into a fully-formatted PowerPoint PPTX file.
                
                 //The JSON structure should include:
                 //  - master_slides: Array of master slide definitions
                 //  - content_slides: Array of content slide definitions
                 //  - Each slide contains shapes with properties (fill, line, shadow, text, etc.)
                
                 //Usage Example:
                 //Uncomment the lines below to convert textbox.json to a PPTX file
                
                 string jsonPath = @"D:\pythonf\c_sharp_project\officeHelperOpenxml\test_ppt\textbox_1_content.json";
                 string outputPptxPath = @"D:\pythonf\c_sharp_project\officeHelperOpenxml\test_ppt\textbox_1_content_json.pptx";
                 return ConvertJsonToPptx(jsonPath, outputPptxPath);
                //
                // Command-line usage:
                //   OfficeHelperOpenXml.exe --input test_ppt\textbox.json --output output.pptx
                //   OfficeHelperOpenXml.exe test_ppt\textbox.json output.pptx
                //
                // For more information, run with --help flag
                // ============================================

                // If no test is enabled, show usage instructions
                ShowUsage();
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 程序执行出错: {ex.Message}");
                Console.WriteLine($"错误详情: {ex.StackTrace}");
                return 1;
            }
        }

        /// <summary>
        /// Parses command-line arguments and executes the appropriate action
        /// </summary>
        /// <param name="args">Command-line arguments</param>
        /// <returns>Exit code</returns>
        private static int ParseCommandLineArguments(string[] args)
        {
            // Check for help flag
            if (args.Length == 1 && (args[0] == "--help" || args[0] == "-h" || args[0] == "/?" || args[0] == "help"))
            {
                ShowCommandLineHelp();
                return 0;
            }

            // Check for compare command
            if (args.Length > 0 && (args[0] == "compare" || args[0] == "--compare"))
            {
                // Remove the "compare" command from args and pass the rest to ComparisonTool
                string[] compareArgs = new string[args.Length - 1];
                Array.Copy(args, 1, compareArgs, 0, args.Length - 1);
                
                var comparisonTool = new ComparisonTool();
                return comparisonTool.Run(compareArgs);
            }

            // Parse --input and --output arguments for conversion
            string inputPath = null;
            string outputPath = null;

            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "--input" || args[i] == "-i") && i + 1 < args.Length)
                {
                    inputPath = args[i + 1];
                    i++; // Skip next argument
                }
                else if ((args[i] == "--output" || args[i] == "-o") && i + 1 < args.Length)
                {
                    outputPath = args[i + 1];
                    i++; // Skip next argument
                }
                else if (!args[i].StartsWith("-"))
                {
                    // Positional arguments: first is input, second is output
                    if (inputPath == null)
                        inputPath = args[i];
                    else if (outputPath == null)
                        outputPath = args[i];
                }
            }

            // Validate arguments
            if (string.IsNullOrEmpty(inputPath) || string.IsNullOrEmpty(outputPath))
            {
                Console.WriteLine("❌ Error: Both input and output paths are required");
                Console.WriteLine();
                ShowCommandLineHelp();
                return 1;
            }

            // Execute conversion
            return ConvertJsonToPptx(inputPath, outputPath);
        }

        /// <summary>
        /// Parse compare command arguments (DISABLED - comparison feature removed)
        /// </summary>
        /// <param name="args">Command-line arguments</param>
        /// <returns>Exit code</returns>
        /*
        private static int ParseCompareCommand(string[] args)
        {
            string generatedPath = null;
            string repairedPath = null;
            string reportPath = null;
            string actionPlanPath = null;

            // Parse arguments
            for (int i = 1; i < args.Length; i++)
            {
                if ((args[i] == "--generated" || args[i] == "-g") && i + 1 < args.Length)
                {
                    generatedPath = args[i + 1];
                    i++;
                }
                else if ((args[i] == "--repaired" || args[i] == "-r") && i + 1 < args.Length)
                {
                    repairedPath = args[i + 1];
                    i++;
                }
                else if ((args[i] == "--report" || args[i] == "-o") && i + 1 < args.Length)
                {
                    reportPath = args[i + 1];
                    i++;
                }
                else if ((args[i] == "--action-plan" || args[i] == "-a") && i + 1 < args.Length)
                {
                    actionPlanPath = args[i + 1];
                    i++;
                }
                else if (!args[i].StartsWith("-"))
                {
                    // Positional arguments
                    if (generatedPath == null)
                        generatedPath = args[i];
                    else if (repairedPath == null)
                        repairedPath = args[i];
                    else if (reportPath == null)
                        reportPath = args[i];
                    else if (actionPlanPath == null)
                        actionPlanPath = args[i];
                }
            }

            // Validate required arguments
            if (string.IsNullOrEmpty(generatedPath) || string.IsNullOrEmpty(repairedPath))
            {
                Console.WriteLine("❌ Error: Both generated and repaired PPTX paths are required");
                Console.WriteLine();
                ShowCompareHelp();
                return 1;
            }

            // Set default output paths if not specified
            if (string.IsNullOrEmpty(reportPath))
            {
                reportPath = "comparison_report.md";
            }

            if (string.IsNullOrEmpty(actionPlanPath))
            {
                actionPlanPath = "action_plan.md";
            }

            // Execute comparison
            return ComparePptxFiles(generatedPath, repairedPath, reportPath, actionPlanPath);
        }
        */

        /// <summary>
        /// Displays command-line help information
        /// </summary>
        private static void ShowCommandLineHelp()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  OfficeHelperOpenXml - Command Line");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Commands:");
            Console.WriteLine("  convert   Convert JSON to PPTX (default)");
            Console.WriteLine("  compare   Compare PPTX files to identify differences");
            Console.WriteLine();
            Console.WriteLine("Convert Usage:");
            Console.WriteLine("  OfficeHelperOpenXml.exe --input <json_file> --output <pptx_file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe -i <json_file> -o <pptx_file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe <json_file> <pptx_file>");
            Console.WriteLine();
            Console.WriteLine("Compare Usage:");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare --content-original <file> --content-generated <file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare --master-original <file> --master-generated <file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare --help");
            Console.WriteLine();
            Console.WriteLine("Convert Options:");
            Console.WriteLine("  --input, -i       Path to the input JSON file");
            Console.WriteLine("  --output, -o      Path to the output PPTX file");
            Console.WriteLine("  --help, -h        Display this help message");
            Console.WriteLine();
            Console.WriteLine("Compare Options:");
            Console.WriteLine("  Run 'OfficeHelperOpenXml.exe compare --help' for detailed comparison options");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  OfficeHelperOpenXml.exe data.json presentation.pptx");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare -co orig.pptx -cg gen.pptx");
            Console.WriteLine();
        }

        /// <summary>
        /// Displays help for compare command (DISABLED - comparison feature removed)
        /// </summary>
        /*
        private static void ShowCompareHelp()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  PPTX Comparison Tool - Command Line");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare --generated <file1> --repaired <file2>");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare -g <file1> -r <file2> -o <report> -a <action_plan>");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare <generated> <repaired> [report] [action_plan]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --generated, -g   Path to the generated PPTX file (required)");
            Console.WriteLine("  --repaired, -r    Path to the repaired PPTX file (required)");
            Console.WriteLine("  --report, -o      Path to save comparison report (default: comparison_report.md)");
            Console.WriteLine("  --action-plan, -a Path to save action plan (default: action_plan.md)");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare generated.pptx repaired.pptx");
            Console.WriteLine("  OfficeHelperOpenXml.exe compare -g gen.pptx -r fixed.pptx -o report.md -a plan.md");
            Console.WriteLine();
        }
        */

        /// <summary>
        /// 处理 PowerPoint 文件
        /// </summary>
        /// <param name="pptPath">PowerPoint 文件路径</param>
        /// <param name="outputPath">输出 JSON 文件路径</param>
        /// <returns>退出代码</returns>
        private static int ProcessPowerPoint(string pptPath, string outputPath)
        {
            // 验证文件是否存在
            if (!File.Exists(pptPath))
            {
                Console.WriteLine($"❌ 错误: 文件不存在 - {pptPath}");
                return 1;
            }

            Console.WriteLine($"📂 开始分析 PowerPoint 文件: {pptPath}");
            Console.WriteLine($"📄 输出文件: {outputPath}");
            Console.WriteLine();

            // 使用 OpenXML SDK 进行分析
            using (var reader = PowerPointReaderFactory.CreateReader(pptPath, out bool success))
            {
                if (!success)
                {
                    Console.WriteLine("❌ 加载 PowerPoint 文件失败！");
                    return 1;
                }

                Console.WriteLine("✅ PowerPoint 文件加载成功！");
                Console.WriteLine();

                // 获取分析结果
                Console.WriteLine("📊 正在分析文件内容...");
                var info = reader.PresentationInfo;
                if (info != null)
                {
                    Console.WriteLine($"📑 幻灯片数量: {info.Slides?.Count ?? 0}");
                    Console.WriteLine($"📏 页面尺寸: {info.SlideWidth} x {info.SlideHeight}");
                }

                // 保存到文件
                Console.WriteLine();
                Console.WriteLine("💾 正在保存分析结果...");
                if (reader.SaveToJson(outputPath))
                {
                    Console.WriteLine($"✅ JSON 文件已保存到: {outputPath}");
                    
                    // 显示文件大小
                    FileInfo fileInfo = new FileInfo(outputPath);
                    Console.WriteLine($"📦 文件大小: {fileInfo.Length / 1024.0:F2} KB");
                }
                else
                {
                    Console.WriteLine("❌ 保存 JSON 文件失败！");
                    return 1;
                }

                Console.WriteLine();
                Console.WriteLine("🎉 分析完成！");
                return 0;
            }
        }

        /// <summary>
        /// 处理 Excel 文件
        /// </summary>
        /// <param name="excelPath">Excel 文件路径</param>
        /// <param name="outputPath">输出 JSON 文件路径</param>
        /// <returns>退出代码</returns>
        private static int ProcessExcel(string excelPath, string outputPath)
        {
            // 验证文件是否存在
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"❌ 错误: 文件不存在 - {excelPath}");
                return 1;
            }

            Console.WriteLine($"📂 开始分析 Excel 文件: {excelPath}");
            Console.WriteLine($"📄 输出文件: {outputPath}");
            Console.WriteLine();

            // 使用 OpenXML SDK 进行分析
            using (var reader = new ExcelReader())
            {
                if (!reader.Load(excelPath))
                {
                    Console.WriteLine("❌ 加载 Excel 文件失败！");
                    return 1;
                }

                Console.WriteLine("✅ Excel 文件加载成功！");
                Console.WriteLine();

                // 获取分析结果
                Console.WriteLine("📊 正在分析文件内容...");
                var sheetNames = reader.GetSheetNames();
                Console.WriteLine($"📋 工作表数量: {sheetNames.Count}");
                Console.WriteLine();

                int totalRows = 0;
                foreach (var sheetName in sheetNames)
                {
                    var data = reader.GetSheetData(sheetName);
                    int rowCount = data.Count;
                    totalRows += rowCount;
                    Console.WriteLine($"  📄 {sheetName}: {rowCount} 行");
                }

                Console.WriteLine();
                Console.WriteLine($"📊 总数据行数: {totalRows}");

                // 保存到文件
                Console.WriteLine();
                Console.WriteLine("💾 正在保存分析结果...");
                var allData = reader.GetAllData();
                var json = JsonConvert.SerializeObject(allData, Formatting.Indented);
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"✅ JSON 文件已保存到: {outputPath}");

                // 显示文件大小
                FileInfo fileInfo = new FileInfo(outputPath);
                Console.WriteLine($"📦 文件大小: {fileInfo.Length / 1024.0:F2} KB");

                Console.WriteLine();
                Console.WriteLine("🎉 Excel 分析完成！");
                return 0;
            }
        }

        /// <summary>
        /// 从 JSON 文件恢复 PowerPoint
        /// </summary>
        /// <param name="jsonPath">输入 JSON 文件路径</param>
        /// <param name="outputPptPath">输出 PPT 文件路径</param>
        /// <returns>退出代码</returns>
        private static int RestorePPTFromJson(string jsonPath, string outputPptPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("    从 JSON 恢复 PowerPoint");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine($"📂 输入 JSON: {jsonPath}");
            Console.WriteLine($"📄 输出 PPT: {outputPptPath}");
            Console.WriteLine();

            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"❌ 错误: JSON 文件不存在 - {jsonPath}");
                return 1;
            }

            try
            {
                // 1. 读取 JSON
                Console.WriteLine("📖 正在读取 JSON 文件...");
                string jsonContent = File.ReadAllText(jsonPath);
                var data = JObject.Parse(jsonContent);

                // 2. 创建 PowerPoint Writer
                Console.WriteLine("📝 正在创建 PowerPoint 文件...");
                using (var writer = new PowerPointWriter())
                {
                    writer.CreateNew();

                    // 3. 处理幻灯片数据
                    var slides = data["Slides"] as JArray;
                    if (slides == null || slides.Count == 0)
                    {
                        Console.WriteLine("⚠️  警告: JSON 中没有幻灯片数据");
                        writer.SaveAs(outputPptPath);
                        return 0;
                    }

                    Console.WriteLine($"📊 找到 {slides.Count} 张幻灯片");
                    Console.WriteLine();

                    // TODO: 实现从 JSON 恢复幻灯片的逻辑
                    // 这里需要根据 PowerPointWriter 的 API 来实现
                    Console.WriteLine("⚠️  注意: JSON 恢复功能需要根据实际 API 实现");

                    // 4. 保存 PPT
                    Console.WriteLine("💾 正在保存 PPT 文件...");
                    writer.SaveAs(outputPptPath);

                    Console.WriteLine();
                    Console.WriteLine($"✅ PPT 已保存: {outputPptPath}");
                    Console.WriteLine();
                    Console.WriteLine("🎉 恢复完成！");

                    return 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 恢复失败: {ex.Message}");
                Console.WriteLine($"错误详情: {ex.StackTrace}");
                return 1;
            }
        }

        /// <summary>
        /// Converts a JSON file to a PowerPoint PPTX file
        /// </summary>
        /// <param name="jsonPath">Path to the input JSON file</param>
        /// <param name="outputPptxPath">Path where the output PPTX file will be saved</param>
        /// <returns>Exit code: 0 for success, non-zero for failure</returns>
        private static int ConvertJsonToPptx(string jsonPath, string outputPptxPath)
        {
            // Validate input parameters
            if (string.IsNullOrWhiteSpace(jsonPath))
            {
                Console.WriteLine("❌ Error: JSON file path cannot be empty");
                return 1;
            }

            if (string.IsNullOrWhiteSpace(outputPptxPath))
            {
                Console.WriteLine("❌ Error: Output PPTX path cannot be empty");
                return 1;
            }

            // Validate JSON file exists
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"❌ Error: JSON file not found - {jsonPath}");
                return 1;
            }

            // Validate output directory
            try
            {
                var outputDir = Path.GetDirectoryName(outputPptxPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Console.WriteLine($"❌ Error: Output directory does not exist - {outputDir}");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error: Invalid output path - {ex.Message}");
                return 1;
            }

            Console.WriteLine("========================================");
            Console.WriteLine("  JSON to PPTX Converter");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine($"📂 Input JSON: {jsonPath}");
            Console.WriteLine($"📄 Output PPTX: {outputPptxPath}");
            Console.WriteLine();

            try
            {
                // Create converter instance
                var converter = new JsonToPptxConverter();

                // Perform conversion
                bool success = converter.Convert(jsonPath, outputPptxPath);

                if (success)
                {
                    Console.WriteLine();
                    Console.WriteLine("========================================");
                    Console.WriteLine("✅ Conversion completed successfully!");
                    Console.WriteLine("========================================");
                    
                    // Display output file info
                    if (File.Exists(outputPptxPath))
                    {
                        FileInfo fileInfo = new FileInfo(outputPptxPath);
                        Console.WriteLine($"📦 Output file size: {fileInfo.Length / 1024.0:F2} KB");
                    }
                    
                    return 0;
                }
                else
                {
                    Console.WriteLine();
                    Console.WriteLine("========================================");
                    Console.WriteLine("❌ Conversion failed");
                    Console.WriteLine("========================================");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("❌ Conversion failed with exception");
                Console.WriteLine("========================================");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return 1;
            }
        }

        /// <summary>
        /// Compare two PPTX files and generate reports (DISABLED - comparison feature removed)
        /// </summary>
        /// <param name="generatedPath">Path to generated PPTX file</param>
        /// <param name="repairedPath">Path to repaired PPTX file</param>
        /// <param name="reportPath">Path to save comparison report</param>
        /// <param name="actionPlanPath">Path to save action plan</param>
        /// <returns>Exit code</returns>
        /*
        private static int ComparePptxFiles(string generatedPath, string repairedPath, string reportPath, string actionPlanPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  PPTX Comparison Tool");
            Console.WriteLine("========================================");
            Console.WriteLine();

            // Validate input files
            if (!File.Exists(generatedPath))
            {
                Console.WriteLine($"❌ Error: Generated PPTX file not found - {generatedPath}");
                return 1;
            }

            if (!File.Exists(repairedPath))
            {
                Console.WriteLine($"❌ Error: Repaired PPTX file not found - {repairedPath}");
                return 1;
            }

            try
            {
                // Create comparison tool
                var comparisonTool = new PptxComparisonTool();

                // Run comparison
                var result = comparisonTool.RunComparison(
                    generatedPath,
                    repairedPath,
                    reportPath,
                    actionPlanPath);

                Console.WriteLine();
                Console.WriteLine("========================================");
                
                if (result.Success)
                {
                    Console.WriteLine("✅ Comparison completed successfully!");
                    Console.WriteLine("========================================");
                    Console.WriteLine();
                    Console.WriteLine("Summary:");
                    Console.WriteLine($"  Total Differences: {result.TotalDifferences}");
                    Console.WriteLine($"  Total Issues: {result.TotalIssues}");
                    Console.WriteLine($"  Generated File Valid: {(result.GeneratedFileValid ? "✓" : "✗")}");
                    Console.WriteLine($"  Repaired File Valid: {(result.RepairedFileValid ? "✓" : "✗")}");
                    Console.WriteLine();
                    Console.WriteLine("Output Files:");
                    Console.WriteLine($"  Report: {result.ReportPath}");
                    Console.WriteLine($"  Action Plan: {result.ActionPlanPath}");
                    
                    return 0;
                }
                else
                {
                    Console.WriteLine("❌ Comparison failed");
                    Console.WriteLine("========================================");
                    Console.WriteLine($"Error: {result.ErrorMessage}");
                    return 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("❌ Comparison failed with exception");
                Console.WriteLine("========================================");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return 1;
            }
        }
        */

        /// <summary>
        /// 显示使用说明
        /// </summary>
        private static void ShowUsage()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  OfficeHelperOpenXml - Usage Guide");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("This program provides Office file analysis and conversion capabilities.");
            Console.WriteLine();
            Console.WriteLine("Available Features:");
            Console.WriteLine("  1. ProcessPowerPoint  - Analyze PowerPoint files and export to JSON");
            Console.WriteLine("  2. ProcessExcel       - Analyze Excel files and export to JSON");
            Console.WriteLine("  3. RestorePPTFromJson - Restore PowerPoint from JSON (legacy)");
            Console.WriteLine("  4. ConvertJsonToPptx  - Convert JSON to PowerPoint PPTX (Recommended)");
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("JSON to PPTX Conversion (Primary Feature)");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Command-line Usage:");
            Console.WriteLine("  OfficeHelperOpenXml.exe --input <json_file> --output <pptx_file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe -i <json_file> -o <pptx_file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe <json_file> <pptx_file>");
            Console.WriteLine("  OfficeHelperOpenXml.exe --help");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("  OfficeHelperOpenXml.exe test_ppt\\textbox.json output.pptx");
            Console.WriteLine();
            Console.WriteLine("Programmatic Usage:");
            Console.WriteLine("  1. Open Program.cs in your editor");
            Console.WriteLine("  2. Locate the Main method");
            Console.WriteLine("  3. Uncomment the ConvertJsonToPptx example code:");
            Console.WriteLine();
            Console.WriteLine("     string jsonPath = @\"test_ppt\\textbox.json\";");
            Console.WriteLine("     string outputPptxPath = @\"test_ppt\\textbox_from_converter.pptx\";");
            Console.WriteLine("     return ConvertJsonToPptx(jsonPath, outputPptxPath);");
            Console.WriteLine();
            Console.WriteLine("  4. Modify file paths as needed");
            Console.WriteLine("  5. Build and run the program");
            Console.WriteLine();
            Console.WriteLine("JSON Structure Requirements:");
            Console.WriteLine("  - master_slides: Array of master slide definitions");
            Console.WriteLine("  - content_slides: Array of content slide definitions");
            Console.WriteLine("  - Each slide contains shapes with properties:");
            Console.WriteLine("    * Position and dimensions (box)");
            Console.WriteLine("    * Fill, line, and shadow properties");
            Console.WriteLine("    * Text content with formatting");
            Console.WriteLine("    * Theme colors and color transforms");
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("Other Features");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("To use other features (PowerPoint/Excel analysis):");
            Console.WriteLine("  1. Open Program.cs");
            Console.WriteLine("  2. Uncomment the desired test code in Main method");
            Console.WriteLine("  3. Update file paths to your test files");
            Console.WriteLine("  4. Build and run");
            Console.WriteLine();
            Console.WriteLine("Library Mode:");
            Console.WriteLine("  To use as a library (DLL) instead of executable:");
            Console.WriteLine("  1. Open OfficeHelperOpenXml.csproj");
            Console.WriteLine("  2. Remove or comment out <OutputType>Exe</OutputType>");
            Console.WriteLine("  3. Rebuild to generate DLL");
            Console.WriteLine();
        }
    }
}
