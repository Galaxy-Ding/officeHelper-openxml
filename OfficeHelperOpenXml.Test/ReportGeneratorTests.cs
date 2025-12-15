using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Core.Comparison;

namespace OfficeHelperOpenXml.Test
{
    public class ReportGeneratorTests : IDisposable
    {
        private readonly string _testOutputDir;
        private readonly ReportGenerator _generator;

        public ReportGeneratorTests()
        {
            _testOutputDir = Path.Combine(Path.GetTempPath(), "ReportGeneratorTests_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(_testOutputDir);
            _generator = new ReportGenerator();
        }

        public void Dispose()
        {
            if (Directory.Exists(_testOutputDir))
            {
                try
                {
                    Directory.Delete(_testOutputDir, true);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        [Fact]
        public void GeneratePairReport_WithValidResult_CreatesReportFile()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();
            string outputPath = Path.Combine(_testOutputDir, "pair_report.md");

            // Act
            string reportPath = _generator.GeneratePairReport(result, outputPath);

            // Assert
            Assert.True(File.Exists(reportPath));
            string content = File.ReadAllText(reportPath);
            Assert.Contains("PPTX Comparison Report: Test Pair", content);
            Assert.Contains("## Summary", content);
            Assert.Contains("Total Differences:", content);
        }

        [Fact]
        public void GeneratePairReport_IncludesTimestamp()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();
            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act
            string reportPath = _generator.GeneratePairReport(result, outputPath);

            // Assert
            string filename = Path.GetFileName(reportPath);
            Assert.Contains("_", filename); // Timestamp separator
            Assert.EndsWith(".md", filename);
        }

        [Fact]
        public void GeneratePairReport_WithMissingFiles_ListsThem()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();
            result.FileComparison.MissingInSecond.Add("ppt/theme/theme1.xml");
            result.FileComparison.MissingInSecond.Add("ppt/slides/_rels/slide1.xml.rels");
            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act
            string reportPath = _generator.GeneratePairReport(result, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("### Missing in Generated", content);
            Assert.Contains("ppt/theme/theme1.xml", content);
            Assert.Contains("ppt/slides/_rels/slide1.xml.rels", content);
        }

        [Fact]
        public void GeneratePairReport_WithExtraFiles_ListsThem()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();
            result.FileComparison.ExtraInSecond.Add("ppt/media/image1.png");
            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act
            string reportPath = _generator.GeneratePairReport(result, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("### Extra in Generated", content);
            Assert.Contains("ppt/media/image1.png", content);
        }

        [Fact]
        public void GeneratePairReport_WithXmlDifferences_GroupsByFile()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();
            
            var xmlResult1 = new XmlComparisonResult { FilePath = "ppt/slides/slide1.xml" };
            xmlResult1.Differences.Add(new XmlDifference
            {
                XPath = "/p:sld/p:cSld/p:spTree/p:sp[1]",
                Type = DifferenceType.MissingElement,
                ExpectedValue = "<p:sp>...</p:sp>",
                ActualValue = "(missing)"
            });
            
            var xmlResult2 = new XmlComparisonResult { FilePath = "ppt/slides/slide2.xml" };
            xmlResult2.Differences.Add(new XmlDifference
            {
                XPath = "/p:sld/@name",
                Type = DifferenceType.DifferentAttributeValue,
                ExpectedValue = "Slide 2",
                ActualValue = "Content Slide"
            });

            result.XmlComparisons["ppt/slides/slide1.xml"] = xmlResult1;
            result.XmlComparisons["ppt/slides/slide2.xml"] = xmlResult2;
            result.TotalDifferences = 2;

            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act
            string reportPath = _generator.GeneratePairReport(result, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("### File: ppt/slides/slide1.xml", content);
            Assert.Contains("### File: ppt/slides/slide2.xml", content);
            Assert.Contains("Missing Element", content);
            Assert.Contains("Different Attribute Value", content);
        }

        [Fact]
        public void GeneratePairReport_WithXmlDifferences_IncludesXPathAndValues()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();
            
            var xmlResult = new XmlComparisonResult { FilePath = "ppt/slides/slide1.xml" };
            xmlResult.Differences.Add(new XmlDifference
            {
                XPath = "/p:sld/p:cSld/@name",
                Type = DifferenceType.DifferentAttributeValue,
                ExpectedValue = "Original Name",
                ActualValue = "Generated Name",
                Description = "Slide name mismatch"
            });

            result.XmlComparisons["ppt/slides/slide1.xml"] = xmlResult;
            result.TotalDifferences = 1;

            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act
            string reportPath = _generator.GeneratePairReport(result, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("/p:sld/p:cSld/@name", content);
            Assert.Contains("Original Name", content);
            Assert.Contains("Generated Name", content);
            Assert.Contains("Slide name mismatch", content);
        }

        [Fact]
        public void GenerateConsolidatedReport_WithBothPairs_CreatesSummary()
        {
            // Arrange
            var summary = new ComparisonSummary
            {
                ContentComparison = CreateSamplePairComparisonResult(),
                MasterComparison = CreateSamplePairComparisonResult()
            };
            summary.ContentComparison.PairName = "Content Pair";
            summary.MasterComparison.PairName = "Master Pair";

            string outputPath = Path.Combine(_testOutputDir, "consolidated.md");

            // Act
            string reportPath = _generator.GenerateConsolidatedReport(summary, outputPath);

            // Assert
            Assert.True(File.Exists(reportPath));
            string content = File.ReadAllText(reportPath);
            Assert.Contains("PPTX Comparison Consolidated Report", content);
            Assert.Contains("## Overall Summary", content);
            Assert.Contains("## Content Comparison", content);
            Assert.Contains("## Master Comparison", content);
        }

        [Fact]
        public void GenerateConsolidatedReport_CalculatesTotalDifferences()
        {
            // Arrange
            var summary = new ComparisonSummary
            {
                ContentComparison = CreateSamplePairComparisonResult(),
                MasterComparison = CreateSamplePairComparisonResult()
            };
            summary.ContentComparison.TotalDifferences = 5;
            summary.MasterComparison.TotalDifferences = 3;

            string outputPath = Path.Combine(_testOutputDir, "consolidated.md");

            // Act
            string reportPath = _generator.GenerateConsolidatedReport(summary, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("Total Differences Across All Pairs: 8", content);
        }

        [Fact]
        public void GenerateRepairPlan_WithDifferences_CreatesActionableSteps()
        {
            // Arrange
            var summary = new ComparisonSummary
            {
                ContentComparison = CreateSamplePairComparisonResult()
            };

            var xmlResult = new XmlComparisonResult { FilePath = "ppt/slides/slide1.xml" };
            xmlResult.Differences.Add(new XmlDifference
            {
                XPath = "/p:sld/p:cSld/p:spTree/p:sp[1]",
                Type = DifferenceType.MissingElement,
                ExpectedValue = "<p:sp>...</p:sp>"
            });

            summary.ContentComparison.XmlComparisons["ppt/slides/slide1.xml"] = xmlResult;
            summary.ContentComparison.TotalDifferences = 1;

            string outputPath = Path.Combine(_testOutputDir, "repair_plan.md");

            // Act
            string reportPath = _generator.GenerateRepairPlan(summary, outputPath);

            // Assert
            Assert.True(File.Exists(reportPath));
            string content = File.ReadAllText(reportPath);
            Assert.Contains("PPTX Comparison Repair Plan", content);
            Assert.Contains("## Repair Actions", content);
            Assert.Contains("**Impact Level:**", content);
            Assert.Contains("**Steps:**", content);
        }

        [Fact]
        public void GenerateRepairPlan_GroupsDifferencesByType()
        {
            // Arrange
            var summary = new ComparisonSummary
            {
                ContentComparison = CreateSamplePairComparisonResult()
            };

            var xmlResult = new XmlComparisonResult { FilePath = "ppt/slides/slide1.xml" };
            xmlResult.Differences.Add(new XmlDifference
            {
                Type = DifferenceType.MissingElement,
                XPath = "/p:sld/p:cSld/p:spTree/p:sp[1]"
            });
            xmlResult.Differences.Add(new XmlDifference
            {
                Type = DifferenceType.MissingElement,
                XPath = "/p:sld/p:cSld/p:spTree/p:sp[2]"
            });
            xmlResult.Differences.Add(new XmlDifference
            {
                Type = DifferenceType.DifferentAttributeValue,
                XPath = "/p:sld/@name"
            });

            summary.ContentComparison.XmlComparisons["ppt/slides/slide1.xml"] = xmlResult;

            string outputPath = Path.Combine(_testOutputDir, "repair_plan.md");

            // Act
            string reportPath = _generator.GenerateRepairPlan(summary, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("Fix Missing Element Issues", content);
            Assert.Contains("Fix Different Attribute Value Issues", content);
        }

        [Fact]
        public void GenerateRepairPlan_IncludesImpactLevels()
        {
            // Arrange
            var summary = new ComparisonSummary
            {
                ContentComparison = CreateSamplePairComparisonResult()
            };

            var xmlResult = new XmlComparisonResult { FilePath = "ppt/slides/slide1.xml" };
            xmlResult.Differences.Add(new XmlDifference
            {
                Type = DifferenceType.MissingElement,
                XPath = "/p:sld/p:cSld/p:spTree/p:sp[1]"
            });

            summary.ContentComparison.XmlComparisons["ppt/slides/slide1.xml"] = xmlResult;

            string outputPath = Path.Combine(_testOutputDir, "repair_plan.md");

            // Act
            string reportPath = _generator.GenerateRepairPlan(summary, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("**Impact Level:** Critical", content);
        }

        [Fact]
        public void GenerateRepairPlan_WithMissingFiles_CreatesRepairAction()
        {
            // Arrange
            var summary = new ComparisonSummary
            {
                ContentComparison = CreateSamplePairComparisonResult()
            };
            summary.ContentComparison.FileComparison.MissingInSecond.Add("ppt/theme/theme1.xml");

            string outputPath = Path.Combine(_testOutputDir, "repair_plan.md");

            // Act
            string reportPath = _generator.GenerateRepairPlan(summary, outputPath);

            // Assert
            string content = File.ReadAllText(reportPath);
            Assert.Contains("Add Missing Files to Generated PPTX", content);
        }

        [Fact]
        public void GeneratePairReport_ThrowsOnNullResult()
        {
            // Arrange
            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _generator.GeneratePairReport(null, outputPath));
        }

        [Fact]
        public void GeneratePairReport_ThrowsOnEmptyOutputPath()
        {
            // Arrange
            var result = CreateSamplePairComparisonResult();

            // Act & Assert
            Assert.Throws<ArgumentException>(() => _generator.GeneratePairReport(result, ""));
        }

        [Fact]
        public void GenerateConsolidatedReport_ThrowsOnNullSummary()
        {
            // Arrange
            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _generator.GenerateConsolidatedReport(null, outputPath));
        }

        [Fact]
        public void GenerateRepairPlan_ThrowsOnNullSummary()
        {
            // Arrange
            string outputPath = Path.Combine(_testOutputDir, "report.md");

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _generator.GenerateRepairPlan(null, outputPath));
        }

        private PairComparisonResult CreateSamplePairComparisonResult()
        {
            return new PairComparisonResult
            {
                PairName = "Test Pair",
                FileComparison = new FileComparisonResult(),
                XmlComparisons = new Dictionary<string, XmlComparisonResult>(),
                TotalDifferences = 0
            };
        }
    }
}
