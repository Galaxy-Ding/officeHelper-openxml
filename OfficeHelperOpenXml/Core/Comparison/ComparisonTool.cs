using System;
using System.IO;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Command-line interface for PPTX comparison tool
    /// </summary>
    public class ComparisonTool
    {
        private readonly ConversionLogger _logger;

        public ComparisonTool()
        {
            _logger = new ConversionLogger();
        }

        /// <summary>
        /// Run the comparison tool with the provided arguments
        /// </summary>
        /// <param name="args">Command-line arguments</param>
        /// <returns>Exit code: 0 for success, non-zero for failure</returns>
        public int Run(string[] args)
        {
            try
            {
                // Check for help flag
                if (args.Length == 0 || IsHelpRequested(args))
                {
                    ShowHelp();
                    return 0;
                }

                // Parse command-line arguments
                ComparisonConfig config = ParseArguments(args);

                // Validate configuration
                if (!ValidateConfig(config))
                {
                    return 1;
                }

                // Display configuration
                DisplayConfiguration(config);

                // Execute comparison
                var orchestrator = new ComparisonOrchestrator();
                ComparisonSummary summary = orchestrator.ExecuteComparison(config);

                // Display results summary
                DisplayResultsSummary(summary);

                return 0;
            }
            catch (ComparisonException ex)
            {
                DisplayError($"Comparison failed: {ex.Message}", ex);
                return 1;
            }
            catch (Exception ex)
            {
                DisplayError($"Unexpected error: {ex.Message}", ex);
                return 1;
            }
        }

        /// <summary>
        /// Check if help is requested
        /// </summary>
        private bool IsHelpRequested(string[] args)
        {
            foreach (var arg in args)
            {
                if (arg == "--help" || arg == "-h" || arg == "/?" || arg == "help")
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Parse command-line arguments into ComparisonConfig
        /// </summary>
        private ComparisonConfig ParseArguments(string[] args)
        {
            var config = new ComparisonConfig();

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                if (arg == "--content-original" || arg == "-co")
                {
                    if (i + 1 < args.Length)
                    {
                        config.ContentOriginalPath = args[++i];
                    }
                    else
                    {
                        throw new ArgumentException($"Missing value for {arg}");
                    }
                }
                else if (arg == "--content-generated" || arg == "-cg")
                {
                    if (i + 1 < args.Length)
                    {
                        config.ContentGeneratedPath = args[++i];
                    }
                    else
                    {
                        throw new ArgumentException($"Missing value for {arg}");
                    }
                }
                else if (arg == "--master-original" || arg == "-mo")
                {
                    if (i + 1 < args.Length)
                    {
                        config.MasterOriginalPath = args[++i];
                    }
                    else
                    {
                        throw new ArgumentException($"Missing value for {arg}");
                    }
                }
                else if (arg == "--master-generated" || arg == "-mg")
                {
                    if (i + 1 < args.Length)
                    {
                        config.MasterGeneratedPath = args[++i];
                    }
                    else
                    {
                        throw new ArgumentException($"Missing value for {arg}");
                    }
                }
                else if (arg == "--output" || arg == "-o")
                {
                    if (i + 1 < args.Length)
                    {
                        config.OutputDirectory = args[++i];
                    }
                    else
                    {
                        throw new ArgumentException($"Missing value for {arg}");
                    }
                }
                else if (!arg.StartsWith("-"))
                {
                    // Positional arguments: content-original, content-generated, master-original, master-generated, output
                    if (string.IsNullOrEmpty(config.ContentOriginalPath))
                    {
                        config.ContentOriginalPath = arg;
                    }
                    else if (string.IsNullOrEmpty(config.ContentGeneratedPath))
                    {
                        config.ContentGeneratedPath = arg;
                    }
                    else if (string.IsNullOrEmpty(config.MasterOriginalPath))
                    {
                        config.MasterOriginalPath = arg;
                    }
                    else if (string.IsNullOrEmpty(config.MasterGeneratedPath))
                    {
                        config.MasterGeneratedPath = arg;
                    }
                    else if (string.IsNullOrEmpty(config.OutputDirectory))
                    {
                        config.OutputDirectory = arg;
                    }
                }
            }

            // Set default output directory if not specified
            if (string.IsNullOrEmpty(config.OutputDirectory))
            {
                config.OutputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "comparison_output");
            }

            return config;
        }

        /// <summary>
        /// Validate the comparison configuration
        /// </summary>
        private bool ValidateConfig(ComparisonConfig config)
        {
            bool isValid = true;

            // Check if at least one pair is specified
            bool hasContentPair = !string.IsNullOrEmpty(config.ContentOriginalPath) && 
                                  !string.IsNullOrEmpty(config.ContentGeneratedPath);
            bool hasMasterPair = !string.IsNullOrEmpty(config.MasterOriginalPath) && 
                                 !string.IsNullOrEmpty(config.MasterGeneratedPath);

            if (!hasContentPair && !hasMasterPair)
            {
                _logger.LogError("Error: At least one comparison pair must be specified");
                _logger.LogError("  Content pair requires: --content-original and --content-generated");
                _logger.LogError("  Master pair requires: --master-original and --master-generated");
                isValid = false;
            }

            // Validate content pair files if specified
            if (hasContentPair)
            {
                if (!File.Exists(config.ContentOriginalPath))
                {
                    _logger.LogError($"Error: Content original file not found: {config.ContentOriginalPath}");
                    isValid = false;
                }
                if (!File.Exists(config.ContentGeneratedPath))
                {
                    _logger.LogError($"Error: Content generated file not found: {config.ContentGeneratedPath}");
                    isValid = false;
                }
            }

            // Validate master pair files if specified
            if (hasMasterPair)
            {
                if (!File.Exists(config.MasterOriginalPath))
                {
                    _logger.LogError($"Error: Master original file not found: {config.MasterOriginalPath}");
                    isValid = false;
                }
                if (!File.Exists(config.MasterGeneratedPath))
                {
                    _logger.LogError($"Error: Master generated file not found: {config.MasterGeneratedPath}");
                    isValid = false;
                }
            }

            // Validate output directory
            if (!string.IsNullOrEmpty(config.OutputDirectory))
            {
                try
                {
                    // Try to create the directory if it doesn't exist
                    if (!Directory.Exists(config.OutputDirectory))
                    {
                        Directory.CreateDirectory(config.OutputDirectory);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error: Cannot create output directory: {ex.Message}");
                    isValid = false;
                }
            }

            return isValid;
        }

        /// <summary>
        /// Display the comparison configuration
        /// </summary>
        private void DisplayConfiguration(ComparisonConfig config)
        {
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("  PPTX Comparison Tool");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Configuration:");

            if (!string.IsNullOrEmpty(config.ContentOriginalPath))
            {
                Console.WriteLine("  Content Pair:");
                Console.WriteLine($"    Original:  {config.ContentOriginalPath}");
                Console.WriteLine($"    Generated: {config.ContentGeneratedPath}");
            }

            if (!string.IsNullOrEmpty(config.MasterOriginalPath))
            {
                Console.WriteLine("  Master Pair:");
                Console.WriteLine($"    Original:  {config.MasterOriginalPath}");
                Console.WriteLine($"    Generated: {config.MasterGeneratedPath}");
            }

            Console.WriteLine($"  Output Directory: {config.OutputDirectory}");
            Console.WriteLine();
        }

        /// <summary>
        /// Display results summary to console
        /// </summary>
        private void DisplayResultsSummary(ComparisonSummary summary)
        {
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("  Comparison Results");
            Console.WriteLine("========================================");
            Console.WriteLine();

            int totalDifferences = 0;
            int totalMissingFiles = 0;
            int totalExtraFiles = 0;
            int totalXmlDifferences = 0;

            if (summary.ContentComparison != null)
            {
                Console.WriteLine("Content Pair:");
                Console.WriteLine($"  Total Differences: {summary.ContentComparison.TotalDifferences}");
                Console.WriteLine($"  Missing Files: {summary.ContentComparison.FileComparison.MissingCount}");
                Console.WriteLine($"  Extra Files: {summary.ContentComparison.FileComparison.ExtraCount}");
                
                int xmlDiffs = 0;
                foreach (var xmlComp in summary.ContentComparison.XmlComparisons.Values)
                {
                    xmlDiffs += xmlComp.Differences.Count;
                }
                Console.WriteLine($"  XML Differences: {xmlDiffs}");
                
                if (!string.IsNullOrEmpty(summary.ContentComparison.ReportPath))
                {
                    Console.WriteLine($"  Report: {summary.ContentComparison.ReportPath}");
                }
                Console.WriteLine();

                totalDifferences += summary.ContentComparison.TotalDifferences;
                totalMissingFiles += summary.ContentComparison.FileComparison.MissingCount;
                totalExtraFiles += summary.ContentComparison.FileComparison.ExtraCount;
                totalXmlDifferences += xmlDiffs;
            }

            if (summary.MasterComparison != null)
            {
                Console.WriteLine("Master Pair:");
                Console.WriteLine($"  Total Differences: {summary.MasterComparison.TotalDifferences}");
                Console.WriteLine($"  Missing Files: {summary.MasterComparison.FileComparison.MissingCount}");
                Console.WriteLine($"  Extra Files: {summary.MasterComparison.FileComparison.ExtraCount}");
                
                int xmlDiffs = 0;
                foreach (var xmlComp in summary.MasterComparison.XmlComparisons.Values)
                {
                    xmlDiffs += xmlComp.Differences.Count;
                }
                Console.WriteLine($"  XML Differences: {xmlDiffs}");
                
                if (!string.IsNullOrEmpty(summary.MasterComparison.ReportPath))
                {
                    Console.WriteLine($"  Report: {summary.MasterComparison.ReportPath}");
                }
                Console.WriteLine();

                totalDifferences += summary.MasterComparison.TotalDifferences;
                totalMissingFiles += summary.MasterComparison.FileComparison.MissingCount;
                totalExtraFiles += summary.MasterComparison.FileComparison.ExtraCount;
                totalXmlDifferences += xmlDiffs;
            }

            Console.WriteLine("Overall Summary:");
            Console.WriteLine($"  Total Differences: {totalDifferences}");
            Console.WriteLine($"  Total Missing Files: {totalMissingFiles}");
            Console.WriteLine($"  Total Extra Files: {totalExtraFiles}");
            Console.WriteLine($"  Total XML Differences: {totalXmlDifferences}");
            Console.WriteLine();

            if (!string.IsNullOrEmpty(summary.ConsolidatedReportPath))
            {
                Console.WriteLine($"Consolidated Report: {summary.ConsolidatedReportPath}");
            }

            if (!string.IsNullOrEmpty(summary.RepairPlanPath))
            {
                Console.WriteLine($"Repair Plan: {summary.RepairPlanPath}");
            }

            Console.WriteLine();
            Console.WriteLine("========================================");
            
            if (totalDifferences == 0)
            {
                _logger.LogSuccess("✓ No differences found - files are identical!");
            }
            else
            {
                _logger.LogWarning($"⚠ Found {totalDifferences} difference(s) - see reports for details");
            }
            
            Console.WriteLine("========================================");
        }

        /// <summary>
        /// Display error message and details
        /// </summary>
        private void DisplayError(string message, Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("  Error");
            Console.WriteLine("========================================");
            Console.WriteLine();
            _logger.LogError(message);
            
            if (ex is ComparisonException compEx)
            {
                if (!string.IsNullOrEmpty(compEx.Context))
                {
                    Console.WriteLine($"Context: {compEx.Context}");
                }
                if (!string.IsNullOrEmpty(compEx.FilePath))
                {
                    Console.WriteLine($"File: {compEx.FilePath}");
                }
                Console.WriteLine($"Category: {compEx.Category}");
            }
            
            Console.WriteLine();
            Console.WriteLine("For more information, run with --help");
            Console.WriteLine("========================================");
        }

        /// <summary>
        /// Show help information
        /// </summary>
        private void ShowHelp()
        {
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("  PPTX Comparison Tool - Help");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Description:");
            Console.WriteLine("  Compare PPTX files to identify differences between original and generated versions.");
            Console.WriteLine("  Supports comparing content slides and master slides separately.");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  ComparisonTool [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --content-original, -co <path>   Path to original content PPTX file");
            Console.WriteLine("  --content-generated, -cg <path>  Path to generated content PPTX file");
            Console.WriteLine("  --master-original, -mo <path>    Path to original master PPTX file");
            Console.WriteLine("  --master-generated, -mg <path>   Path to generated master PPTX file");
            Console.WriteLine("  --output, -o <path>              Output directory for reports (default: ./comparison_output)");
            Console.WriteLine("  --help, -h                       Show this help message");
            Console.WriteLine();
            Console.WriteLine("Positional Arguments:");
            Console.WriteLine("  You can also provide paths in order without flags:");
            Console.WriteLine("  ComparisonTool <content-orig> <content-gen> <master-orig> <master-gen> [output]");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine();
            Console.WriteLine("  Compare content pair only:");
            Console.WriteLine("    ComparisonTool -co original.pptx -cg generated.pptx");
            Console.WriteLine();
            Console.WriteLine("  Compare both content and master pairs:");
            Console.WriteLine("    ComparisonTool -co content_orig.pptx -cg content_gen.pptx \\");
            Console.WriteLine("                   -mo master_orig.pptx -mg master_gen.pptx");
            Console.WriteLine();
            Console.WriteLine("  Using positional arguments:");
            Console.WriteLine("    ComparisonTool content_orig.pptx content_gen.pptx \\");
            Console.WriteLine("                   master_orig.pptx master_gen.pptx output_dir");
            Console.WriteLine();
            Console.WriteLine("  Specify custom output directory:");
            Console.WriteLine("    ComparisonTool -co orig.pptx -cg gen.pptx -o ./reports");
            Console.WriteLine();
            Console.WriteLine("Output:");
            Console.WriteLine("  The tool generates the following reports in the output directory:");
            Console.WriteLine("  - content_pair_report.md      Detailed comparison of content files");
            Console.WriteLine("  - master_pair_report.md       Detailed comparison of master files");
            Console.WriteLine("  - consolidated_report.md      Summary of all comparisons");
            Console.WriteLine("  - repair_plan.md              Suggested fixes for identified issues");
            Console.WriteLine();
            Console.WriteLine("Requirements:");
            Console.WriteLine("  - At least one comparison pair (content or master) must be specified");
            Console.WriteLine("  - All specified PPTX files must exist");
            Console.WriteLine("  - Output directory must be writable");
            Console.WriteLine();
            Console.WriteLine("========================================");
        }
    }
}
