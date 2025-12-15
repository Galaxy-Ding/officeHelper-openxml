using System.Collections.Generic;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Result of extracting a PPTX file
    /// </summary>
    public class ExtractionResult
    {
        public bool Success { get; set; }
        public string ExtractedPath { get; set; }
        public List<string> Files { get; set; }
        public string ErrorMessage { get; set; }

        public ExtractionResult()
        {
            Files = new List<string>();
        }
    }

    /// <summary>
    /// Result of comparing two file lists
    /// </summary>
    public class FileComparisonResult
    {
        public List<string> MissingInSecond { get; set; }  // In first but not in second
        public List<string> ExtraInSecond { get; set; }    // In second but not in first
        public List<string> CommonFiles { get; set; }      // In both

        public int MissingCount => MissingInSecond.Count;
        public int ExtraCount => ExtraInSecond.Count;
        public int CommonCount => CommonFiles.Count;

        public FileComparisonResult()
        {
            MissingInSecond = new List<string>();
            ExtraInSecond = new List<string>();
            CommonFiles = new List<string>();
        }
    }

    /// <summary>
    /// Result of comparing two XML files
    /// </summary>
    public class XmlComparisonResult
    {
        public List<XmlDifference> Differences { get; set; }
        public bool AreIdentical => Differences.Count == 0;
        public string FilePath { get; set; }

        public XmlComparisonResult()
        {
            Differences = new List<XmlDifference>();
        }
    }

    /// <summary>
    /// Configuration for comparison process
    /// </summary>
    public class ComparisonConfig
    {
        public string ContentOriginalPath { get; set; }
        public string ContentGeneratedPath { get; set; }
        public string MasterOriginalPath { get; set; }
        public string MasterGeneratedPath { get; set; }
        public string OutputDirectory { get; set; }
    }

    /// <summary>
    /// Summary of entire comparison process
    /// </summary>
    public class ComparisonSummary
    {
        public PairComparisonResult ContentComparison { get; set; }
        public PairComparisonResult MasterComparison { get; set; }
        public string ConsolidatedReportPath { get; set; }
        public string RepairPlanPath { get; set; }
    }

    /// <summary>
    /// Result of comparing a single pair of PPTX files
    /// </summary>
    public class PairComparisonResult
    {
        public string PairName { get; set; }
        public FileComparisonResult FileComparison { get; set; }
        public Dictionary<string, XmlComparisonResult> XmlComparisons { get; set; }
        public string ReportPath { get; set; }
        public int TotalDifferences { get; set; }

        public PairComparisonResult()
        {
            XmlComparisons = new Dictionary<string, XmlComparisonResult>();
        }
    }

    /// <summary>
    /// Result of cleanup operations
    /// </summary>
    public class CleanupResult
    {
        public int TotalDirectories { get; set; }
        public int SuccessfulCleanups { get; set; }
        public int FailedCleanups { get; set; }
        public List<string> Errors { get; set; }

        public CleanupResult()
        {
            Errors = new List<string>();
        }
    }
}
