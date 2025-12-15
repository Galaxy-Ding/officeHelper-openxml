using System.Collections.Generic;
using System.Linq;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Compares two file lists and identifies differences using set operations
    /// </summary>
    public class FileComparator
    {
        /// <summary>
        /// Compare two file lists and identify missing, extra, and common files
        /// </summary>
        /// <param name="files1">First file list (original)</param>
        /// <param name="files2">Second file list (generated)</param>
        /// <returns>FileComparisonResult containing set differences and intersection</returns>
        public FileComparisonResult Compare(List<string> files1, List<string> files2)
        {
            var result = new FileComparisonResult();

            // Handle null inputs
            if (files1 == null)
            {
                files1 = new List<string>();
            }
            if (files2 == null)
            {
                files2 = new List<string>();
            }

            // Convert to HashSet for efficient set operations
            var set1 = new HashSet<string>(files1);
            var set2 = new HashSet<string>(files2);

            // Set difference: files1 - files2 (in first but not in second)
            result.MissingInSecond = set1.Except(set2).OrderBy(f => f).ToList();

            // Set difference: files2 - files1 (in second but not in first)
            result.ExtraInSecond = set2.Except(set1).OrderBy(f => f).ToList();

            // Set intersection: files1 ∩ files2 (in both)
            result.CommonFiles = set1.Intersect(set2).OrderBy(f => f).ToList();

            return result;
        }
    }
}
