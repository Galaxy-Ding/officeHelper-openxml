using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Orchestrates the entire PPTX comparison workflow
    /// Coordinates extraction, comparison, reporting, and cleanup
    /// </summary>
    public class ComparisonOrchestrator
    {
        private readonly PptxExtractor _extractor;
        private readonly FileComparator _fileComparator;
        private readonly XmlComparator _xmlComparator;
        private readonly ReportGenerator _reportGenerator;
        private readonly TempDirManager _tempDirManager;
        private readonly ConversionLogger _logger;

        public ComparisonOrchestrator()
        {
            _logger = new ConversionLogger();
            _tempDirManager = new TempDirManager(_logger);
            _extractor = new PptxExtractor(_tempDirManager, _logger);
            _fileComparator = new FileComparator();
            _xmlComparator = new XmlComparator();
            _reportGenerator = new ReportGenerator(_logger);
        }

        public ComparisonOrchestrator(
            PptxExtractor extractor,
            FileComparator fileComparator,
            XmlComparator xmlComparator,
            ReportGenerator reportGenerator,
            TempDirManager tempDirManager,
            ConversionLogger logger)
        {
            _extractor = extractor ?? throw new ArgumentNullException(nameof(extractor));
            _fileComparator = fileComparator ?? throw new ArgumentNullException(nameof(fileComparator));
            _xmlComparator = xmlComparator ?? throw new ArgumentNullException(nameof(xmlComparator));
            _reportGenerator = reportGenerator ?? throw new ArgumentNullException(nameof(reportGenerator));
            _tempDirManager = tempDirManager ?? throw new ArgumentNullException(nameof(tempDirManager));
            _logger = logger ?? new ConversionLogger();
        }

        /// <summary>
        /// Execute the full comparison workflow for both content and master pairs
        /// </summary>
        /// <param name="config">Configuration containing file paths and output directory</param>
        /// <returns>ComparisonSummary containing results from all comparisons</returns>
        public ComparisonSummary ExecuteComparison(ComparisonConfig config)
        {
            if (config == null)
                throw new ArgumentNullException(nameof(config));

            _logger.LogInfo("=== Starting PPTX Comparison Workflow ===");
            _logger.LogInfo($"Output Directory: {config.OutputDirectory}");

            var summary = new ComparisonSummary();

            try
            {
                // Ensure output directory exists
                if (!string.IsNullOrWhiteSpace(config.OutputDirectory))
                {
                    Directory.CreateDirectory(config.OutputDirectory);
                }

                // Phase 1: Content Comparison
                if (!string.IsNullOrWhiteSpace(config.ContentOriginalPath) && 
                    !string.IsNullOrWhiteSpace(config.ContentGeneratedPath))
                {
                    _logger.LogInfo("");
                    _logger.LogInfo("=== Phase 1: Content Comparison ===");
                    summary.ContentComparison = ComparePair(
                        config.ContentOriginalPath,
                        config.ContentGeneratedPath,
                        "Content Pair",
                        config.OutputDirectory);
                }

                // Phase 2: Master Comparison
                if (!string.IsNullOrWhiteSpace(config.MasterOriginalPath) && 
                    !string.IsNullOrWhiteSpace(config.MasterGeneratedPath))
                {
                    _logger.LogInfo("");
                    _logger.LogInfo("=== Phase 2: Master Comparison ===");
                    summary.MasterComparison = ComparePair(
                        config.MasterOriginalPath,
                        config.MasterGeneratedPath,
                        "Master Pair",
                        config.OutputDirectory);
                }

                // Phase 3: Generate Consolidated Reports
                _logger.LogInfo("");
                _logger.LogInfo("=== Phase 3: Generating Consolidated Reports ===");

                if (!string.IsNullOrWhiteSpace(config.OutputDirectory))
                {
                    // Generate consolidated summary report
                    string consolidatedReportPath = Path.Combine(config.OutputDirectory, "consolidated_report.md");
                    summary.ConsolidatedReportPath = _reportGenerator.GenerateConsolidatedReport(
                        summary, 
                        consolidatedReportPath);

                    // Generate repair plan
                    string repairPlanPath = Path.Combine(config.OutputDirectory, "repair_plan.md");
                    summary.RepairPlanPath = _reportGenerator.GenerateRepairPlan(
                        summary, 
                        repairPlanPath);
                }

                _logger.LogInfo("");
                _logger.LogSuccess("=== Comparison Workflow Completed Successfully ===");
                LogSummaryStatistics(summary);

                return summary;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Comparison workflow failed: {ex.Message}");
                throw;
            }
            finally
            {
                // Phase 4: Cleanup - Always attempt cleanup even if errors occurred
                _logger.LogInfo("");
                _logger.LogInfo("=== Phase 4: Cleanup ===");
                CleanupTempDirectories();
            }
        }

        /// <summary>
        /// Compare a single pair of PPTX files (original vs generated)
        /// </summary>
        /// <param name="originalPath">Path to original PPTX file</param>
        /// <param name="generatedPath">Path to generated PPTX file</param>
        /// <param name="pairName">Name for this comparison pair</param>
        /// <param name="outputDirectory">Directory for output reports</param>
        /// <returns>PairComparisonResult containing all comparison results</returns>
        private PairComparisonResult ComparePair(
            string originalPath, 
            string generatedPath, 
            string pairName,
            string outputDirectory)
        {
            _logger.LogInfo($"Comparing pair: {pairName}");
            _logger.LogInfo($"  Original:  {originalPath}");
            _logger.LogInfo($"  Generated: {generatedPath}");

            var result = new PairComparisonResult
            {
                PairName = pairName
            };

            try
            {
                // Step 1: Extract both PPTX files
                _logger.LogInfo("Step 1: Extracting PPTX files...");
                ExtractionResult originalExtraction = _extractor.Extract(originalPath);
                ExtractionResult generatedExtraction = _extractor.Extract(generatedPath);

                if (!originalExtraction.Success)
                {
                    throw new ComparisonException($"Failed to extract original PPTX: {originalExtraction.ErrorMessage}")
                    {
                        Context = "ComparePair - Original Extraction",
                        FilePath = originalPath,
                        Category = ErrorCategory.Archive
                    };
                }

                if (!generatedExtraction.Success)
                {
                    throw new ComparisonException($"Failed to extract generated PPTX: {generatedExtraction.ErrorMessage}")
                    {
                        Context = "ComparePair - Generated Extraction",
                        FilePath = generatedPath,
                        Category = ErrorCategory.Archive
                    };
                }

                _logger.LogSuccess($"Extracted {originalExtraction.Files.Count} files from original");
                _logger.LogSuccess($"Extracted {generatedExtraction.Files.Count} files from generated");

                // Step 2: Compare file lists
                _logger.LogInfo("Step 2: Comparing file lists...");
                result.FileComparison = _fileComparator.Compare(
                    originalExtraction.Files, 
                    generatedExtraction.Files);

                _logger.LogInfo($"  Missing in generated: {result.FileComparison.MissingCount}");
                _logger.LogInfo($"  Extra in generated: {result.FileComparison.ExtraCount}");
                _logger.LogInfo($"  Common files: {result.FileComparison.CommonCount}");

                // Step 3: Compare XML content for common files
                _logger.LogInfo("Step 3: Comparing XML content...");
                int xmlDifferenceCount = 0;

                foreach (var commonFile in result.FileComparison.CommonFiles)
                {
                    // Only compare XML and .rels files
                    if (commonFile.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                        commonFile.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    {
                        string originalFilePath = Path.Combine(originalExtraction.ExtractedPath, commonFile);
                        string generatedFilePath = Path.Combine(generatedExtraction.ExtractedPath, commonFile);

                        try
                        {
                            XmlComparisonResult xmlResult = _xmlComparator.Compare(
                                originalFilePath, 
                                generatedFilePath);

                            xmlResult.FilePath = commonFile;
                            result.XmlComparisons[commonFile] = xmlResult;

                            if (xmlResult.Differences.Count > 0)
                            {
                                xmlDifferenceCount += xmlResult.Differences.Count;
                                _logger.LogWarning($"  {commonFile}: {xmlResult.Differences.Count} difference(s)");
                            }
                        }
                        catch (ComparisonException ex)
                        {
                            _logger.LogWarning($"  Failed to compare {commonFile}: {ex.Message}");
                            // Continue with other files even if one fails
                        }
                    }
                }

                _logger.LogInfo($"Total XML differences found: {xmlDifferenceCount}");

                // Calculate total differences
                result.TotalDifferences = 
                    result.FileComparison.MissingCount + 
                    result.FileComparison.ExtraCount + 
                    xmlDifferenceCount;

                // Step 4: Generate pair report
                if (!string.IsNullOrWhiteSpace(outputDirectory))
                {
                    _logger.LogInfo("Step 4: Generating pair report...");
                    string reportFileName = $"{pairName.Replace(" ", "_").ToLower()}_report.md";
                    string reportPath = Path.Combine(outputDirectory, reportFileName);
                    result.ReportPath = _reportGenerator.GeneratePairReport(result, reportPath);
                }

                _logger.LogSuccess($"Pair comparison completed: {pairName}");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to compare pair '{pairName}': {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Clean up all temporary directories
        /// </summary>
        private void CleanupTempDirectories()
        {
            try
            {
                CleanupResult cleanupResult = _tempDirManager.CleanupAll();

                if (cleanupResult.FailedCleanups > 0)
                {
                    _logger.LogWarning($"Cleanup completed with {cleanupResult.FailedCleanups} failures");
                    foreach (var error in cleanupResult.Errors)
                    {
                        _logger.LogWarning($"  {error}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log cleanup errors but don't throw - cleanup failures should not fail the overall process
                _logger.LogWarning($"Cleanup encountered an error: {ex.Message}");
            }
        }

        /// <summary>
        /// Log summary statistics for the entire comparison
        /// </summary>
        private void LogSummaryStatistics(ComparisonSummary summary)
        {
            _logger.LogInfo("");
            _logger.LogInfo("=== Summary Statistics ===");

            int totalDifferences = 0;
            int totalMissingFiles = 0;
            int totalExtraFiles = 0;
            int totalXmlDifferences = 0;

            if (summary.ContentComparison != null)
            {
                totalDifferences += summary.ContentComparison.TotalDifferences;
                totalMissingFiles += summary.ContentComparison.FileComparison.MissingCount;
                totalExtraFiles += summary.ContentComparison.FileComparison.ExtraCount;
                totalXmlDifferences += summary.ContentComparison.XmlComparisons.Values.Sum(x => x.Differences.Count);
            }

            if (summary.MasterComparison != null)
            {
                totalDifferences += summary.MasterComparison.TotalDifferences;
                totalMissingFiles += summary.MasterComparison.FileComparison.MissingCount;
                totalExtraFiles += summary.MasterComparison.FileComparison.ExtraCount;
                totalXmlDifferences += summary.MasterComparison.XmlComparisons.Values.Sum(x => x.Differences.Count);
            }

            _logger.LogInfo($"Total Differences: {totalDifferences}");
            _logger.LogInfo($"  Missing Files: {totalMissingFiles}");
            _logger.LogInfo($"  Extra Files: {totalExtraFiles}");
            _logger.LogInfo($"  XML Differences: {totalXmlDifferences}");

            if (!string.IsNullOrWhiteSpace(summary.ConsolidatedReportPath))
            {
                _logger.LogInfo($"Consolidated Report: {summary.ConsolidatedReportPath}");
            }

            if (!string.IsNullOrWhiteSpace(summary.RepairPlanPath))
            {
                _logger.LogInfo($"Repair Plan: {summary.RepairPlanPath}");
            }
        }
    }
}
