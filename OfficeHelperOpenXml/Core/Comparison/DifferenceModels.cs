using System.Collections.Generic;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Represents a single difference found in XML comparison
    /// </summary>
    public class XmlDifference
    {
        public string XPath { get; set; }
        public DifferenceType Type { get; set; }
        public string ExpectedValue { get; set; }
        public string ActualValue { get; set; }
        public string Description { get; set; }
    }

    /// <summary>
    /// Types of differences that can be found in XML comparison
    /// </summary>
    public enum DifferenceType
    {
        MissingElement,
        ExtraElement,
        MissingAttribute,
        ExtraAttribute,
        DifferentAttributeValue,
        DifferentTextContent
    }

    /// <summary>
    /// Category of differences with impact assessment
    /// </summary>
    public class DifferenceCategory
    {
        public string Name { get; set; }
        public ImpactLevel Impact { get; set; }
        public List<XmlDifference> Differences { get; set; }
        public string CodeLocation { get; set; }
        public string RepairInstructions { get; set; }

        public DifferenceCategory()
        {
            Differences = new List<XmlDifference>();
        }
    }

    /// <summary>
    /// Impact level of a difference
    /// </summary>
    public enum ImpactLevel
    {
        Critical,   // Prevents file from opening
        Important,  // Causes visible issues
        Minor       // Cosmetic differences
    }

    /// <summary>
    /// Repair action to fix identified differences
    /// </summary>
    public class RepairAction
    {
        public int Priority { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public ImpactLevel Impact { get; set; }
        public string AffectedFiles { get; set; }
        public string CodeLocation { get; set; }
        public List<string> Steps { get; set; }
        public List<XmlDifference> RelatedDifferences { get; set; }

        public RepairAction()
        {
            Steps = new List<string>();
            RelatedDifferences = new List<XmlDifference>();
        }
    }

    /// <summary>
    /// Exception thrown during comparison operations
    /// </summary>
    public class ComparisonException : System.Exception
    {
        public string Context { get; set; }
        public string FilePath { get; set; }
        public ErrorCategory Category { get; set; }

        public ComparisonException(string message) : base(message)
        {
        }

        public ComparisonException(string message, System.Exception innerException) 
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// Category of errors that can occur during comparison
    /// </summary>
    public enum ErrorCategory
    {
        FileSystem,
        Archive,
        Xml,
        System
    }
}
