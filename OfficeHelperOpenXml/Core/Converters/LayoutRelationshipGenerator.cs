using System;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// Generates relationship files for slide layouts
    /// </summary>
    public class LayoutRelationshipGenerator
    {
        /// <summary>
        /// Generates relationship files for a slide layout
        /// </summary>
        /// <param name="layoutPart">The slide layout part to generate relationships for</param>
        /// <param name="masterPart">The parent slide master part</param>
        /// <param name="relationshipIdManager">The relationship ID manager for sequential IDs</param>
        public void GenerateLayoutRelationships(SlideLayoutPart layoutPart, SlideMasterPart masterPart, RelationshipIdManager relationshipIdManager)
        {
            if (layoutPart == null)
            {
                throw new ArgumentNullException(nameof(layoutPart));
            }

            if (masterPart == null)
            {
                throw new ArgumentNullException(nameof(masterPart));
            }

            if (relationshipIdManager == null)
            {
                throw new ArgumentNullException(nameof(relationshipIdManager));
            }

            // Reset relationship counter for layout's .rels file
            relationshipIdManager.ResetCounter();

            // Check if relationship already exists
            var existingRelationshipId = layoutPart.GetIdOfPart(masterPart);
            if (string.IsNullOrEmpty(existingRelationshipId))
            {
                // Add relationship from layout to master slide using sequential ID
                // This creates the .rels file in ppt/slideLayouts/_rels/
                layoutPart.AddPart(masterPart, relationshipIdManager.GetNextId());
            }
        }
    }
}
