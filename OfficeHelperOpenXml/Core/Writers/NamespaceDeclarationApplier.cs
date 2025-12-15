using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Core.Writers
{
    /// <summary>
    /// Utility class for applying namespace declarations to OpenXML elements.
    /// Ensures consistent namespace declarations across all generated XML files.
    /// </summary>
    public static class NamespaceDeclarationApplier
    {
        // Namespace constants
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";

        /// <summary>
        /// Applies required namespace declarations to a Presentation element.
        /// </summary>
        /// <param name="presentation">The Presentation element to apply namespaces to</param>
        public static void ApplyPresentationNamespaces(Presentation presentation)
        {
            if (presentation == null)
                return;

            AddNamespaceIfMissing(presentation, "a", DrawingNamespace);
            AddNamespaceIfMissing(presentation, "r", RelationshipsNamespace);
            AddNamespaceIfMissing(presentation, "p", PresentationNamespace);
        }

        /// <summary>
        /// Applies required namespace declarations to a Slide element.
        /// </summary>
        /// <param name="slide">The Slide element to apply namespaces to</param>
        public static void ApplySlideNamespaces(Slide slide)
        {
            if (slide == null)
                return;

            AddNamespaceIfMissing(slide, "a", DrawingNamespace);
            AddNamespaceIfMissing(slide, "r", RelationshipsNamespace);
            AddNamespaceIfMissing(slide, "p", PresentationNamespace);
        }

        /// <summary>
        /// Applies required namespace declarations to a SlideMaster element.
        /// </summary>
        /// <param name="slideMaster">The SlideMaster element to apply namespaces to</param>
        public static void ApplySlideMasterNamespaces(SlideMaster slideMaster)
        {
            if (slideMaster == null)
                return;

            AddNamespaceIfMissing(slideMaster, "a", DrawingNamespace);
            AddNamespaceIfMissing(slideMaster, "r", RelationshipsNamespace);
            AddNamespaceIfMissing(slideMaster, "p", PresentationNamespace);
        }

        /// <summary>
        /// Applies required namespace declarations to a SlideLayout element.
        /// </summary>
        /// <param name="slideLayout">The SlideLayout element to apply namespaces to</param>
        public static void ApplySlideLayoutNamespaces(SlideLayout slideLayout)
        {
            if (slideLayout == null)
                return;

            AddNamespaceIfMissing(slideLayout, "a", DrawingNamespace);
            AddNamespaceIfMissing(slideLayout, "r", RelationshipsNamespace);
            AddNamespaceIfMissing(slideLayout, "p", PresentationNamespace);
        }

        /// <summary>
        /// Applies required namespace declarations to a PresentationProperties element.
        /// </summary>
        /// <param name="presentationProperties">The PresentationProperties element to apply namespaces to</param>
        public static void ApplyPresPropsNamespaces(PresentationProperties presentationProperties)
        {
            if (presentationProperties == null)
                return;

            AddNamespaceIfMissing(presentationProperties, "a", DrawingNamespace);
            AddNamespaceIfMissing(presentationProperties, "r", RelationshipsNamespace);
        }

        /// <summary>
        /// Helper method to add a namespace declaration if it doesn't already exist.
        /// </summary>
        /// <param name="element">The OpenXmlElement to add the namespace to</param>
        /// <param name="prefix">The namespace prefix (e.g., "a", "r", "p")</param>
        /// <param name="uri">The namespace URI</param>
        private static void AddNamespaceIfMissing(OpenXmlElement element, string prefix, string uri)
        {
            if (element == null)
                return;

            // Check if namespace already exists
            var existingNamespace = element.NamespaceDeclarations
                .FirstOrDefault(ns => ns.Key == prefix);

            // Only add if not already present
            if (existingNamespace.Key == null)
            {
                element.AddNamespaceDeclaration(prefix, uri);
            }
        }
    }
}
