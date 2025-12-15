using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// Manages sequential relationship ID generation and relative path calculation
    /// </summary>
    public class RelationshipIdManager
    {
        private int _counter = 0;

        /// <summary>
        /// Gets the next sequential relationship ID
        /// </summary>
        /// <returns>Relationship ID in format rId1, rId2, rId3, etc.</returns>
        public string GetNextId()
        {
            _counter++;
            return $"rId{_counter}";
        }

        /// <summary>
        /// Resets the counter for a new relationship file
        /// </summary>
        public void ResetCounter()
        {
            _counter = 0;
        }

        /// <summary>
        /// Calculates relative path from one location to another
        /// </summary>
        /// <param name="from">Source path (e.g., "ppt/slideMasters/_rels/")</param>
        /// <param name="to">Target path (e.g., "ppt/theme/theme1.xml")</param>
        /// <returns>Relative path (e.g., "../theme/theme1.xml")</returns>
        public static string CalculateRelativePath(string from, string to)
        {
            // Remove file name from 'from' path if present
            if (from.EndsWith(".xml.rels"))
            {
                from = Path.GetDirectoryName(from);
            }

            // Split paths into segments
            var fromParts = from.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var toParts = to.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);

            // Find common prefix length
            int commonLength = 0;
            for (int i = 0; i < Math.Min(fromParts.Length, toParts.Length); i++)
            {
                if (fromParts[i] == toParts[i])
                    commonLength++;
                else
                    break;
            }

            // Build relative path
            var relativeParts = new List<string>();

            // Add ../ for each level up
            for (int i = commonLength; i < fromParts.Length; i++)
            {
                relativeParts.Add("..");
            }

            // Add remaining target path segments
            for (int i = commonLength; i < toParts.Length; i++)
            {
                relativeParts.Add(toParts[i]);
            }

            return string.Join("/", relativeParts);
        }
    }
}
