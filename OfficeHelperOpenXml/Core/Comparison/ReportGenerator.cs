using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Generates markdown reports for PPTX comparison results
    /// </summary>
    public class ReportGenerator
    {
        private readonly ConversionLogger _logger;

        public ReportGenerator()
        {
            _logger = new ConversionLogger();
        }

        public ReportGenerator(ConversionLogger logger)
        {
            _logger = logger ?? new ConversionLogger();
        }

        /// <summary>
        /// Generate comparison report for a single pair of PPTX files
        /// </summary>
        /// <param name="result">The comparison result for the pair</param>
        /// <param name="outputPath">Path where the report should be saved</param>
        /// <returns>Path to the generated report file</returns>
        public string GeneratePairReport(PairComparisonResult result, string outputPath)
        {
            if (result == null)
                throw new ArgumentNullException(nameof(result));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));

            _logger.LogInfo($"Generating pair report for: {result.PairName}");

            var sb = new StringBuilder();

            // Header
            sb.AppendLine($"# PPTX Comparison Report: {result.PairName}");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            // Summary section
            sb.AppendLine("## Summary");
            int totalXmlDifferences = result.XmlComparisons.Values.Sum(x => x.Differences.Count);
            sb.AppendLine($"- Total Differences: {result.TotalDifferences}");
            sb.AppendLine($"- Missing Files: {result.FileComparison.MissingCount}");
            sb.AppendLine($"- Extra Files: {result.FileComparison.ExtraCount}");
            sb.AppendLine($"- XML Differences: {totalXmlDifferences}");
            sb.AppendLine();

            // File-level differences
            sb.AppendLine("## File-Level Differences");
            sb.AppendLine();

            if (result.FileComparison.MissingCount > 0)
            {
                sb.AppendLine("### Missing in Generated");
                foreach (var file in result.FileComparison.MissingInSecond)
                {
                    sb.AppendLine($"- {file}");
                }
                sb.AppendLine();
            }

            if (result.FileComparison.ExtraCount > 0)
            {
                sb.AppendLine("### Extra in Generated");
                foreach (var file in result.FileComparison.ExtraInSecond)
                {
                    sb.AppendLine($"- {file}");
                }
                sb.AppendLine();
            }

            if (result.FileComparison.MissingCount == 0 && result.FileComparison.ExtraCount == 0)
            {
                sb.AppendLine("No file-level differences found.");
                sb.AppendLine();
            }

            // XML differences grouped by file
            if (totalXmlDifferences > 0)
            {
                sb.AppendLine("## XML Differences");
                sb.AppendLine();

                // Group by file path
                foreach (var kvp in result.XmlComparisons.OrderBy(x => x.Key))
                {
                    if (kvp.Value.Differences.Count > 0)
                    {
                        sb.AppendLine($"### File: {kvp.Key}");
                        sb.AppendLine();

                        int diffIndex = 1;
                        foreach (var diff in kvp.Value.Differences)
                        {
                            sb.AppendLine($"{diffIndex}. **{FormatDifferenceType(diff.Type)}** at `{diff.XPath}`");
                            
                            if (!string.IsNullOrEmpty(diff.ExpectedValue))
                            {
                                sb.AppendLine($"   - Expected: `{EscapeMarkdown(diff.ExpectedValue)}`");
                            }
                            
                            if (!string.IsNullOrEmpty(diff.ActualValue))
                            {
                                sb.AppendLine($"   - Actual: `{EscapeMarkdown(diff.ActualValue)}`");
                            }
                            
                            if (!string.IsNullOrEmpty(diff.Description))
                            {
                                sb.AppendLine($"   - Description: {diff.Description}");
                            }
                            
                            sb.AppendLine();
                            diffIndex++;
                        }
                    }
                }
            }
            else
            {
                sb.AppendLine("## XML Differences");
                sb.AppendLine();
                sb.AppendLine("No XML differences found.");
                sb.AppendLine();
            }

            // Write to file
            string reportPath = AddTimestampToFilename(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            File.WriteAllText(reportPath, sb.ToString());

            _logger.LogSuccess($"Pair report generated: {reportPath}");
            return reportPath;
        }

        /// <summary>
        /// Generate consolidated summary report across multiple pairs
        /// </summary>
        /// <param name="summary">The comparison summary containing all pair results</param>
        /// <param name="outputPath">Path where the report should be saved</param>
        /// <returns>Path to the generated report file</returns>
        public string GenerateConsolidatedReport(ComparisonSummary summary, string outputPath)
        {
            if (summary == null)
                throw new ArgumentNullException(nameof(summary));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));

            _logger.LogInfo("Generating consolidated report");

            var sb = new StringBuilder();

            // Header
            sb.AppendLine("# PPTX Comparison Consolidated Report");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            // Overall summary
            sb.AppendLine("## Overall Summary");
            sb.AppendLine();

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

            sb.AppendLine($"- Total Differences Across All Pairs: {totalDifferences}");
            sb.AppendLine($"- Total Missing Files: {totalMissingFiles}");
            sb.AppendLine($"- Total Extra Files: {totalExtraFiles}");
            sb.AppendLine($"- Total XML Differences: {totalXmlDifferences}");
            sb.AppendLine();

            // Content comparison summary
            if (summary.ContentComparison != null)
            {
                sb.AppendLine("## Content Comparison");
                AppendPairSummary(sb, summary.ContentComparison);
            }

            // Master comparison summary
            if (summary.MasterComparison != null)
            {
                sb.AppendLine("## Master Comparison");
                AppendPairSummary(sb, summary.MasterComparison);
            }

            // Links to detailed reports
            sb.AppendLine("## Detailed Reports");
            sb.AppendLine();
            if (summary.ContentComparison != null && !string.IsNullOrEmpty(summary.ContentComparison.ReportPath))
            {
                sb.AppendLine($"- [Content Comparison Report]({Path.GetFileName(summary.ContentComparison.ReportPath)})");
            }
            if (summary.MasterComparison != null && !string.IsNullOrEmpty(summary.MasterComparison.ReportPath))
            {
                sb.AppendLine($"- [Master Comparison Report]({Path.GetFileName(summary.MasterComparison.ReportPath)})");
            }
            if (!string.IsNullOrEmpty(summary.RepairPlanPath))
            {
                sb.AppendLine($"- [Repair Plan]({Path.GetFileName(summary.RepairPlanPath)})");
            }
            sb.AppendLine();

            // Write to file
            string reportPath = AddTimestampToFilename(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            File.WriteAllText(reportPath, sb.ToString());

            _logger.LogSuccess($"Consolidated report generated: {reportPath}");
            return reportPath;
        }

        /// <summary>
        /// Generate repair plan with actionable instructions
        /// </summary>
        /// <param name="summary">The comparison summary containing all differences</param>
        /// <param name="outputPath">Path where the repair plan should be saved</param>
        /// <returns>Path to the generated repair plan file</returns>
        public string GenerateRepairPlan(ComparisonSummary summary, string outputPath)
        {
            if (summary == null)
                throw new ArgumentNullException(nameof(summary));
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));

            _logger.LogInfo("Generating repair plan");

            var sb = new StringBuilder();

            // Header
            sb.AppendLine("# PPTX Comparison Repair Plan");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            sb.AppendLine("## Overview");
            sb.AppendLine();
            sb.AppendLine("This document provides a prioritized repair plan to address differences found between");
            sb.AppendLine("original PPTX files and their JSON-generated counterparts.");
            sb.AppendLine();

            // Collect all differences
            var allDifferences = new List<XmlDifference>();
            var filePathMap = new Dictionary<XmlDifference, string>();

            if (summary.ContentComparison != null)
            {
                foreach (var kvp in summary.ContentComparison.XmlComparisons)
                {
                    foreach (var diff in kvp.Value.Differences)
                    {
                        allDifferences.Add(diff);
                        filePathMap[diff] = kvp.Key;
                    }
                }
            }

            if (summary.MasterComparison != null)
            {
                foreach (var kvp in summary.MasterComparison.XmlComparisons)
                {
                    foreach (var diff in kvp.Value.Differences)
                    {
                        allDifferences.Add(diff);
                        filePathMap[diff] = kvp.Key;
                    }
                }
            }

            // Categorize and generate repair actions
            var repairActions = CategorizeAndGenerateRepairActions(allDifferences, filePathMap, summary);

            // Sort by priority
            repairActions = repairActions.OrderBy(a => a.Priority).ToList();

            // Summary statistics
            sb.AppendLine("## Summary");
            sb.AppendLine();
            sb.AppendLine($"- Total Repair Actions: {repairActions.Count}");
            sb.AppendLine($"- Critical: {repairActions.Count(a => a.Impact == ImpactLevel.Critical)}");
            sb.AppendLine($"- Important: {repairActions.Count(a => a.Impact == ImpactLevel.Important)}");
            sb.AppendLine($"- Minor: {repairActions.Count(a => a.Impact == ImpactLevel.Minor)}");
            sb.AppendLine();

            // Repair actions by priority
            sb.AppendLine("## Repair Actions");
            sb.AppendLine();

            foreach (var action in repairActions)
            {
                sb.AppendLine($"### {action.Priority}. {action.Title}");
                sb.AppendLine();
                sb.AppendLine($"**Impact Level:** {action.Impact}");
                sb.AppendLine();
                sb.AppendLine($"**Description:** {action.Description}");
                sb.AppendLine();
                
                if (!string.IsNullOrEmpty(action.AffectedFiles))
                {
                    sb.AppendLine($"**Affected Files:** {action.AffectedFiles}");
                    sb.AppendLine();
                }
                
                if (!string.IsNullOrEmpty(action.CodeLocation))
                {
                    sb.AppendLine($"**Code Location:** `{action.CodeLocation}`");
                    sb.AppendLine();
                }
                
                if (action.Steps.Count > 0)
                {
                    sb.AppendLine("**Steps:**");
                    foreach (var step in action.Steps)
                    {
                        sb.AppendLine($"1. {step}");
                    }
                    sb.AppendLine();
                }
                
                if (action.RelatedDifferences.Count > 0)
                {
                    sb.AppendLine($"**Related Differences:** {action.RelatedDifferences.Count} difference(s)");
                    sb.AppendLine();
                }
            }

            // Write to file
            string reportPath = AddTimestampToFilename(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath));
            File.WriteAllText(reportPath, sb.ToString());

            _logger.LogSuccess($"Repair plan generated: {reportPath}");
            return reportPath;
        }

        /// <summary>
        /// Append summary information for a pair comparison
        /// </summary>
        private void AppendPairSummary(StringBuilder sb, PairComparisonResult result)
        {
            sb.AppendLine();
            sb.AppendLine($"**Pair Name:** {result.PairName}");
            sb.AppendLine();
            
            int xmlDiffCount = result.XmlComparisons.Values.Sum(x => x.Differences.Count);
            
            sb.AppendLine($"- Total Differences: {result.TotalDifferences}");
            sb.AppendLine($"- Missing Files: {result.FileComparison.MissingCount}");
            sb.AppendLine($"- Extra Files: {result.FileComparison.ExtraCount}");
            sb.AppendLine($"- XML Differences: {xmlDiffCount}");
            sb.AppendLine();
        }

        /// <summary>
        /// Categorize differences and generate repair actions
        /// </summary>
        private List<RepairAction> CategorizeAndGenerateRepairActions(
            List<XmlDifference> differences, 
            Dictionary<XmlDifference, string> filePathMap,
            ComparisonSummary summary)
        {
            var actions = new List<RepairAction>();
            int priority = 1;

            // Analyze patterns in differences to create more specific repair actions
            var categorizedActions = AnalyzeDifferencePatterns(differences, filePathMap, summary);
            
            // Add pattern-based actions first (they are more specific)
            foreach (var action in categorizedActions.OrderBy(a => a.Priority))
            {
                action.Priority = priority++;
                actions.Add(action);
            }

            // Group differences by type for generic repair actions
            var missingElementDiffs = differences.Where(d => d.Type == DifferenceType.MissingElement).ToList();
            var differentAttributeDiffs = differences.Where(d => d.Type == DifferenceType.DifferentAttributeValue).ToList();
            
            // Add generic repair actions for difference types
            if (missingElementDiffs.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = priority++,
                    Title = "Fix Missing Element Issues",
                    Description = $"Found {missingElementDiffs.Count} missing XML elements in generated PPTX",
                    Impact = ImpactLevel.Important,
                    CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                    RelatedDifferences = missingElementDiffs,
                    Steps = new List<string>
                    {
                        "Review missing elements in detailed comparison report",
                        "Identify which components are responsible for creating these elements",
                        "Add missing element generation logic",
                        "Test that elements are properly created"
                    }
                });
            }
            
            if (differentAttributeDiffs.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = priority++,
                    Title = "Fix Different Attribute Value Issues",
                    Description = $"Found {differentAttributeDiffs.Count} attribute value mismatches in generated PPTX",
                    Impact = ImpactLevel.Important,
                    CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                    RelatedDifferences = differentAttributeDiffs,
                    Steps = new List<string>
                    {
                        "Review attribute differences in detailed comparison report",
                        "Identify which components set these attributes",
                        "Update attribute value generation logic",
                        "Test that attributes match expected values"
                    }
                });
            }

            // Add file-level repair actions
            int missingFileCount = 0;
            int extraFileCount = 0;
            var missingFiles = new List<string>();
            var extraFiles = new List<string>();

            if (summary.ContentComparison != null)
            {
                missingFileCount += summary.ContentComparison.FileComparison.MissingCount;
                extraFileCount += summary.ContentComparison.FileComparison.ExtraCount;
                missingFiles.AddRange(summary.ContentComparison.FileComparison.MissingInSecond);
                extraFiles.AddRange(summary.ContentComparison.FileComparison.ExtraInSecond);
            }

            if (summary.MasterComparison != null)
            {
                missingFileCount += summary.MasterComparison.FileComparison.MissingCount;
                extraFileCount += summary.MasterComparison.FileComparison.ExtraCount;
                missingFiles.AddRange(summary.MasterComparison.FileComparison.MissingInSecond);
                extraFiles.AddRange(summary.MasterComparison.FileComparison.ExtraInSecond);
            }

            if (missingFileCount > 0)
            {
                // Add generic missing files action first
                actions.Add(new RepairAction
                {
                    Priority = priority++,
                    Title = "Add Missing Files to Generated PPTX",
                    Description = $"The generated PPTX is missing {missingFileCount} file(s) present in the original",
                    Impact = ImpactLevel.Critical,
                    CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                    AffectedFiles = string.Join(", ", missingFiles),
                    Steps = new List<string>
                    {
                        "Review list of missing files",
                        "Implement generation logic for each missing file type",
                        "Update content types and relationships",
                        "Test file integrity"
                    }
                });
            }

            if (missingFileCount > 0)
            {
                // Categorize missing files
                var metadataFiles = missingFiles.Where(f => f.StartsWith("docProps") || 
                    f.Contains("presProps") || f.Contains("viewProps") || f.Contains("tableStyles")).ToList();
                var themeFiles = missingFiles.Where(f => f.Contains("theme") && !f.Contains("slideMasters")).ToList();
                var otherFiles = missingFiles.Except(metadataFiles).Except(themeFiles).ToList();

                if (metadataFiles.Count > 0)
                {
                    actions.Add(new RepairAction
                    {
                        Priority = priority++,
                        Title = "Add Missing Package Metadata Files",
                        Description = $"The generated PPTX is missing {metadataFiles.Count} metadata file(s) required by OOXML specification: {string.Join(", ", metadataFiles.Select(Path.GetFileName))}",
                        Impact = ImpactLevel.Critical,
                        CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                        Steps = new List<string>
                        {
                            "Create GeneratePackageMetadata() method in JsonToPptxConverter",
                            "Implement docProps/app.xml generation with basic application properties",
                            "Implement docProps/core.xml generation with creator, dates, and revision info",
                            "Implement ppt/presProps.xml with default presentation properties",
                            "Implement ppt/viewProps.xml with default view settings",
                            "Implement ppt/tableStyles.xml with default table styles",
                            "Update [Content_Types].xml to register these files",
                            "Update _rels/.rels to add relationships to docProps files",
                            "Test that generated PPTX opens correctly in PowerPoint"
                        }
                    });
                }

                if (themeFiles.Count > 0)
                {
                    actions.Add(new RepairAction
                    {
                        Priority = priority++,
                        Title = "Fix Theme File Location",
                        Description = $"Theme files should be at ppt/theme/ level, not under slideMasters: {string.Join(", ", themeFiles)}",
                        Impact = ImpactLevel.Important,
                        CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                        Steps = new List<string>
                        {
                            "Locate theme file generation code",
                            "Change output path from ppt/slideMasters/theme/ to ppt/theme/",
                            "Update relationship references in presentation.xml.rels",
                            "Update [Content_Types].xml entries",
                            "Test that theme is properly applied"
                        }
                    });
                }

                if (otherFiles.Count > 0)
                {
                    actions.Add(new RepairAction
                    {
                        Priority = priority++,
                        Title = "Add Other Missing Files",
                        Description = $"Additional missing files: {string.Join(", ", otherFiles)}",
                        Impact = ImpactLevel.Important,
                        CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                        Steps = new List<string>
                        {
                            "Review each missing file type",
                            "Implement generation logic for each file type",
                            "Update content types and relationships",
                            "Test file integrity"
                        }
                    });
                }
            }

            if (extraFileCount > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = priority++,
                    Title = "Remove Extra Files from Generated PPTX",
                    Description = $"The generated PPTX contains {extraFileCount} extra file(s) not present in the original: {string.Join(", ", extraFiles.Select(Path.GetFileName))}",
                    Impact = ImpactLevel.Minor,
                    CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                    Steps = new List<string>
                    {
                        "Review why extra masters/layouts are being created (slideLayout2, slideMaster2)",
                        "Review why duplicate theme files exist (theme1.xml, theme2.xml)",
                        "Remove unnecessary file generation logic",
                        "Update [Content_Types].xml to remove entries for deleted files",
                        "Update relationship files to remove references",
                        "Verify the PPTX still opens correctly after removal"
                    }
                });
            }

            return actions;
        }

        /// <summary>
        /// Analyze difference patterns to create specific repair actions
        /// </summary>
        private List<RepairAction> AnalyzeDifferencePatterns(
            List<XmlDifference> differences,
            Dictionary<XmlDifference, string> filePathMap,
            ComparisonSummary summary)
        {
            var actions = new List<RepairAction>();

            // Pattern 1: Missing namespace declarations at root level
            var namespaceIssues = differences.Where(d => 
                d.Type == DifferenceType.MissingAttribute &&
                (d.XPath.Contains("@{http://www.w3.org/2000/xmlns/}a") || 
                 d.XPath.Contains("@{http://www.w3.org/2000/xmlns/}r")) &&
                (d.XPath.StartsWith("/presentation") || d.XPath.StartsWith("/sld") || 
                 d.XPath.StartsWith("/sldMaster") || d.XPath.StartsWith("/sldLayout"))).ToList();

            if (namespaceIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 1,
                    Title = "Fix Missing Namespace Declarations at Root Elements",
                    Description = $"Found {namespaceIssues.Count} missing namespace declarations (xmlns:a, xmlns:r) at root elements. These should be declared at the root level, not on child elements.",
                    Impact = ImpactLevel.Critical,
                    AffectedFiles = string.Join(", ", namespaceIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "OfficeHelperOpenXml.Core.Writers.SlideWriter, JsonToPptxConverter",
                    RelatedDifferences = namespaceIssues,
                    Steps = new List<string>
                    {
                        "Locate methods creating root elements: presentation, sld, sldMaster, sldLayout",
                        "Add xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" to root element",
                        "Add xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" to root element",
                        "Remove xmlns declarations from child elements (they inherit from root)",
                        "Test XML parsing and validation"
                    }
                });
            }

            // Pattern 2: Relationship ID format (Guid vs rId format)
            var relationshipIdIssues = differences.Where(d =>
                d.Type == DifferenceType.DifferentAttributeValue &&
                d.XPath.Contains("/@Id") &&
                d.ExpectedValue != null && d.ExpectedValue.StartsWith("rId") &&
                d.ActualValue != null && d.ActualValue.StartsWith("R") && d.ActualValue.Length > 10).ToList();

            if (relationshipIdIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 2,
                    Title = "Fix Relationship ID Format (Use rId1, rId2 instead of GUIDs)",
                    Description = $"Found {relationshipIdIssues.Count} relationship IDs using Guid format instead of simple rId format. Original uses rId1, rId2, etc., but generated uses R{{Guid}}.",
                    Impact = ImpactLevel.Important,
                    AffectedFiles = string.Join(", ", relationshipIdIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "Relationship creation methods throughout codebase",
                    RelatedDifferences = relationshipIdIssues,
                    Steps = new List<string>
                    {
                        "Search for relationship ID generation code (likely uses Guid.NewGuid())",
                        "Replace with sequential counter: rId1, rId2, rId3, etc.",
                        "Maintain separate counter for each .rels file",
                        "Ensure IDs are unique within each relationship file",
                        "Update all relationship references to use new format",
                        "Test that all relationships resolve correctly"
                    }
                });
            }

            // Pattern 3: Relationship path format (absolute vs relative)
            var pathFormatIssues = differences.Where(d =>
                d.Type == DifferenceType.DifferentAttributeValue &&
                d.XPath.Contains("/@Target") &&
                d.ExpectedValue != null && d.ExpectedValue.Contains("../") &&
                d.ActualValue != null && d.ActualValue.StartsWith("/ppt/")).ToList();

            if (pathFormatIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 2,
                    Title = "Fix Relationship Paths (Use Relative Paths)",
                    Description = $"Found {pathFormatIssues.Count} relationship paths using absolute format instead of relative. Should use ../ notation for parent directories.",
                    Impact = ImpactLevel.Important,
                    AffectedFiles = string.Join(", ", pathFormatIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "Relationship Target attribute setting in SlideWriter or JsonToPptxConverter",
                    RelatedDifferences = pathFormatIssues,
                    Steps = new List<string>
                    {
                        "Locate relationship Target attribute setting code",
                        "Implement relative path calculation based on source and target file locations",
                        "Use ../ notation to navigate to parent directories",
                        "Example: from ppt/slideMasters/_rels/ to ppt/theme/ should be ../theme/theme1.xml",
                        "Test that all relationships resolve correctly",
                        "Verify files open in PowerPoint"
                    }
                });
            }

            // Pattern 4: Missing grpSpPr/xfrm children
            var missingTransformChildren = differences.Where(d =>
                d.Type == DifferenceType.MissingElement &&
                d.XPath.Contains("/grpSpPr[1]/xfrm[1]/") &&
                (d.XPath.Contains("/off[1]") || d.XPath.Contains("/ext[1]") || 
                 d.XPath.Contains("/chOff[1]") || d.XPath.Contains("/chExt[1]"))).ToList();

            if (missingTransformChildren.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 2,
                    Title = "Add Missing Group Shape Transform Children",
                    Description = $"Found {missingTransformChildren.Count} missing child elements in grpSpPr/xfrm. Need to add: off, ext, chOff, chExt elements.",
                    Impact = ImpactLevel.Important,
                    AffectedFiles = string.Join(", ", missingTransformChildren.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "Shape tree creation in SlideWriter or shape component classes",
                    RelatedDifferences = missingTransformChildren,
                    Steps = new List<string>
                    {
                        "Locate grpSpPr/xfrm creation code in shape tree generation",
                        "Add <off x=\"0\" y=\"0\"/> element",
                        "Add <ext cx=\"0\" cy=\"0\"/> element",
                        "Add <chOff x=\"0\" y=\"0\"/> element",
                        "Add <chExt cx=\"0\" cy=\"0\"/> element (use slide dimensions)",
                        "Test that shapes render correctly with proper positioning"
                    }
                });
            }

            // Pattern 5: Text run property issues
            var textRunIssues = differences.Where(d =>
                d.XPath.Contains("/rPr[1]/@") &&
                (d.XPath.Contains("lang") || d.XPath.Contains("altLang") || 
                 d.XPath.Contains("dirty") || d.XPath.Contains("smtClean") || d.XPath.Contains("sz"))).ToList();

            if (textRunIssues.Count > 0)
            {
                var missingLangAttrs = textRunIssues.Where(d => d.Type == DifferenceType.MissingAttribute).ToList();
                var extraSizeAttrs = textRunIssues.Where(d => d.Type == DifferenceType.ExtraAttribute && d.XPath.Contains("sz")).ToList();

                if (missingLangAttrs.Count > 0 || extraSizeAttrs.Count > 0)
                {
                    actions.Add(new RepairAction
                    {
                        Priority = 2,
                        Title = "Fix Text Run Properties (Language and Formatting)",
                        Description = $"Found {textRunIssues.Count} text run property issues. Missing language attributes (lang, altLang, dirty, smtClean) and extra explicit formatting (sz) that should inherit from theme.",
                        Impact = ImpactLevel.Important,
                        AffectedFiles = string.Join(", ", textRunIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                        CodeLocation = "OfficeHelperOpenXml.Components.TextComponent",
                        RelatedDifferences = textRunIssues,
                        Steps = new List<string>
                        {
                            "Locate text run property (rPr) creation in TextComponent",
                            "Add lang attribute (e.g., lang=\"zh-CN\" for Chinese)",
                            "Add altLang attribute (e.g., altLang=\"en-US\")",
                            "Add dirty=\"0\" and smtClean=\"0\" attributes",
                            "Remove explicit sz (font size) attribute - let it inherit from theme",
                            "Remove explicit latin font specification - let it inherit",
                            "Remove explicit solidFill - let it inherit from theme",
                            "Test that text renders correctly with inherited styles"
                        }
                    });
                }
            }

            // Pattern 6: Shape property issues (txBox, wrap, rtlCol, etc.)
            var shapePropertyIssues = differences.Where(d =>
                (d.XPath.Contains("/cNvSpPr[1]/@txBox") || 
                 d.XPath.Contains("/bodyPr[1]/@wrap") ||
                 d.XPath.Contains("/bodyPr[1]/@rtlCol") ||
                 d.XPath.Contains("/@preserve") ||
                 d.XPath.Contains("/@userDrawn"))).ToList();

            if (shapePropertyIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 2,
                    Title = "Fix Shape Property Attributes",
                    Description = $"Found {shapePropertyIssues.Count} missing shape property attributes: txBox (text box indicator), wrap, rtlCol, preserve, userDrawn.",
                    Impact = ImpactLevel.Important,
                    AffectedFiles = string.Join(", ", shapePropertyIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "Elements classes (TextBoxElement, AutoShapeElement), TextComponent",
                    RelatedDifferences = shapePropertyIssues,
                    Steps = new List<string>
                    {
                        "Locate shape creation code in Elements classes",
                        "Add txBox=\"1\" attribute to cNvSpPr for text box shapes",
                        "Add wrap=\"none\" attribute to bodyPr for text wrapping",
                        "Add rtlCol=\"0\" attribute to bodyPr for text direction",
                        "Add preserve=\"1\" to sldLayout where appropriate",
                        "Add userDrawn=\"1\" to sldLayout and nvPr for user-drawn shapes",
                        "Test that shapes behave correctly (editing, resizing, etc.)"
                    }
                });
            }

            // Pattern 7: Extra elements that shouldn't be there
            var extraElementIssues = differences.Where(d =>
                d.Type == DifferenceType.ExtraElement &&
                (d.XPath.Contains("/spLocks[1]") || d.XPath.Contains("/ph[1]") || 
                 d.XPath.Contains("/ln[1]") || d.XPath.Contains("/pPr[1]"))).ToList();

            if (extraElementIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 2,
                    Title = "Remove Extra Elements (spLocks, ph, ln, pPr)",
                    Description = $"Found {extraElementIssues.Count} extra elements being created that don't exist in original: spLocks (shape locks), ph (placeholder), ln (line), pPr (paragraph properties).",
                    Impact = ImpactLevel.Important,
                    AffectedFiles = string.Join(", ", extraElementIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "Shape creation in Elements classes, TextComponent",
                    RelatedDifferences = extraElementIssues,
                    Steps = new List<string>
                    {
                        "Locate spLocks element creation - remove or make conditional",
                        "Locate ph (placeholder) element creation - only add for actual placeholders",
                        "Locate ln (line) element creation - remove if not needed",
                        "Locate pPr element creation - only add when paragraph properties differ from default",
                        "Test that shapes still function correctly without these elements"
                    }
                });
            }

            // Pattern 8: Missing text style definitions
            var textStyleIssues = differences.Where(d =>
                d.Type == DifferenceType.MissingElement &&
                (d.XPath.Contains("/defaultTextStyle[1]") || 
                 d.XPath.Contains("/txStyles[1]") ||
                 d.XPath.Contains("/endParaRPr[1]"))).ToList();

            if (textStyleIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 2,
                    Title = "Add Missing Text Style Definitions",
                    Description = $"Found {textStyleIssues.Count} missing text style elements: defaultTextStyle (presentation level), txStyles (master level), endParaRPr (paragraph level).",
                    Impact = ImpactLevel.Important,
                    AffectedFiles = string.Join(", ", textStyleIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "JsonToPptxConverter (presentation), SlideWriter (master), TextComponent (paragraph)",
                    RelatedDifferences = textStyleIssues,
                    Steps = new List<string>
                    {
                        "Add defaultTextStyle element to presentation.xml",
                        "Add txStyles element to slide master",
                        "Add endParaRPr element at end of each paragraph",
                        "Copy style definitions from original files as templates",
                        "Test that text renders with correct default styles"
                    }
                });
            }

            // Pattern 9: Missing geometry adjustment list
            var geometryIssues = differences.Where(d =>
                d.Type == DifferenceType.MissingElement &&
                d.XPath.Contains("/prstGeom[1]/avLst[1]")).ToList();

            if (geometryIssues.Count > 0)
            {
                actions.Add(new RepairAction
                {
                    Priority = 3,
                    Title = "Add Missing Geometry Adjustment Lists",
                    Description = $"Found {geometryIssues.Count} missing avLst (adjustment value list) elements in preset geometry definitions.",
                    Impact = ImpactLevel.Minor,
                    AffectedFiles = string.Join(", ", geometryIssues.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "Shape geometry creation code",
                    RelatedDifferences = geometryIssues,
                    Steps = new List<string>
                    {
                        "Locate prstGeom creation code",
                        "Add empty <avLst/> element after prst attribute",
                        "Test that shapes render correctly"
                    }
                });
            }

            // Group remaining differences by type for generic actions
            var remainingDifferences = differences.Except(
                namespaceIssues.Concat(relationshipIdIssues).Concat(pathFormatIssues)
                .Concat(missingTransformChildren).Concat(textRunIssues).Concat(shapePropertyIssues)
                .Concat(extraElementIssues).Concat(textStyleIssues).Concat(geometryIssues)).ToList();

            var groupedByType = remainingDifferences.GroupBy(d => d.Type);

            foreach (var group in groupedByType)
            {
                var action = new RepairAction
                {
                    Priority = 3,
                    Title = $"Fix Remaining {FormatDifferenceType(group.Key)} Issues",
                    Description = GenerateDescriptionForType(group.Key, group.Count()),
                    Impact = DetermineImpact(group.Key),
                    AffectedFiles = string.Join(", ", group.Select(d => filePathMap.ContainsKey(d) ? filePathMap[d] : "Unknown").Distinct()),
                    CodeLocation = "OfficeHelperOpenXml.Core.Converters.JsonToPptxConverter",
                    RelatedDifferences = group.ToList()
                };

                action.Steps = GenerateStepsForType(group.Key);
                actions.Add(action);
            }

            return actions;
        }

        /// <summary>
        /// Determine impact level based on difference type
        /// </summary>
        private ImpactLevel DetermineImpact(DifferenceType type)
        {
            switch (type)
            {
                case DifferenceType.MissingElement:
                    return ImpactLevel.Critical;
                case DifferenceType.ExtraElement:
                    return ImpactLevel.Important;
                case DifferenceType.MissingAttribute:
                    return ImpactLevel.Important;
                case DifferenceType.ExtraAttribute:
                    return ImpactLevel.Minor;
                case DifferenceType.DifferentAttributeValue:
                    return ImpactLevel.Important;
                case DifferenceType.DifferentTextContent:
                    return ImpactLevel.Important;
                default:
                    return ImpactLevel.Minor;
            }
        }

        /// <summary>
        /// Generate description for a difference type
        /// </summary>
        private string GenerateDescriptionForType(DifferenceType type, int count)
        {
            switch (type)
            {
                case DifferenceType.MissingElement:
                    return $"Found {count} missing element(s) in the generated PPTX. These elements exist in the original but are not being created by the converter.";
                case DifferenceType.ExtraElement:
                    return $"Found {count} extra element(s) in the generated PPTX. These elements are being created but don't exist in the original.";
                case DifferenceType.MissingAttribute:
                    return $"Found {count} missing attribute(s) in the generated PPTX. These attributes should be added to the corresponding elements.";
                case DifferenceType.ExtraAttribute:
                    return $"Found {count} extra attribute(s) in the generated PPTX. These attributes are being added but don't exist in the original.";
                case DifferenceType.DifferentAttributeValue:
                    return $"Found {count} attribute(s) with different values. The converter is setting incorrect values for these attributes.";
                case DifferenceType.DifferentTextContent:
                    return $"Found {count} text content difference(s). The converter is generating different text content than expected.";
                default:
                    return $"Found {count} difference(s) of type {type}.";
            }
        }

        /// <summary>
        /// Generate repair steps for a difference type
        /// </summary>
        private List<string> GenerateStepsForType(DifferenceType type)
        {
            var steps = new List<string>();

            switch (type)
            {
                case DifferenceType.MissingElement:
                    steps.Add("Identify the XML element that should be created");
                    steps.Add("Locate the converter method responsible for this element type");
                    steps.Add("Add code to create the missing element with appropriate attributes");
                    steps.Add("Test to verify the element is now present in generated files");
                    break;

                case DifferenceType.ExtraElement:
                    steps.Add("Identify the XML element that should not be created");
                    steps.Add("Locate the converter method creating this element");
                    steps.Add("Remove or conditionally skip the element creation logic");
                    steps.Add("Test to verify the element is no longer present");
                    break;

                case DifferenceType.MissingAttribute:
                    steps.Add("Identify the attribute that should be added");
                    steps.Add("Locate the code creating the parent element");
                    steps.Add("Add code to set the missing attribute with the correct value");
                    steps.Add("Test to verify the attribute is now present");
                    break;

                case DifferenceType.ExtraAttribute:
                    steps.Add("Identify the attribute that should not be added");
                    steps.Add("Locate the code setting this attribute");
                    steps.Add("Remove the attribute setting logic");
                    steps.Add("Test to verify the attribute is no longer present");
                    break;

                case DifferenceType.DifferentAttributeValue:
                    steps.Add("Compare the expected and actual attribute values");
                    steps.Add("Locate the code setting this attribute value");
                    steps.Add("Fix the logic to generate the correct value");
                    steps.Add("Test to verify the attribute now has the correct value");
                    break;

                case DifferenceType.DifferentTextContent:
                    steps.Add("Compare the expected and actual text content");
                    steps.Add("Locate the code setting this text content");
                    steps.Add("Fix the text generation logic");
                    steps.Add("Test to verify the text content is now correct");
                    break;
            }

            return steps;
        }

        /// <summary>
        /// Format difference type for display
        /// </summary>
        private string FormatDifferenceType(DifferenceType type)
        {
            switch (type)
            {
                case DifferenceType.MissingElement:
                    return "Missing Element";
                case DifferenceType.ExtraElement:
                    return "Extra Element";
                case DifferenceType.MissingAttribute:
                    return "Missing Attribute";
                case DifferenceType.ExtraAttribute:
                    return "Extra Attribute";
                case DifferenceType.DifferentAttributeValue:
                    return "Different Attribute Value";
                case DifferenceType.DifferentTextContent:
                    return "Different Text Content";
                default:
                    return type.ToString();
            }
        }

        /// <summary>
        /// Escape markdown special characters
        /// </summary>
        private string EscapeMarkdown(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // Escape backticks and other markdown characters
            return text.Replace("`", "\\`")
                      .Replace("*", "\\*")
                      .Replace("_", "\\_")
                      .Replace("[", "\\[")
                      .Replace("]", "\\]");
        }

        /// <summary>
        /// Add timestamp to filename before extension
        /// </summary>
        private string AddTimestampToFilename(string path)
        {
            string directory = Path.GetDirectoryName(path);
            string filename = Path.GetFileNameWithoutExtension(path);
            string extension = Path.GetExtension(path);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            return Path.Combine(directory, $"{filename}_{timestamp}{extension}");
        }
    }
}
