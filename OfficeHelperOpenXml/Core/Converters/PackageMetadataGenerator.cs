using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Core.Converters
{
    /// <summary>
    /// Generates required package metadata files for PPTX documents
    /// </summary>
    public class PackageMetadataGenerator
    {
        /// <summary>
        /// Generates all required metadata files for a presentation document
        /// </summary>
        /// <param name="document">The presentation document to add metadata to</param>
        /// <param name="relationshipIdManager">The relationship ID manager for sequential IDs</param>
        public void GenerateAllMetadata(PresentationDocument document, RelationshipIdManager relationshipIdManager)
        {
            GenerateAppXml(document, relationshipIdManager);
            GenerateCoreXml(document);
            GeneratePresProps(document, relationshipIdManager);
            GenerateViewProps(document, relationshipIdManager);
            GenerateTableStyles(document, relationshipIdManager);
        }

        /// <summary>
        /// Generates docProps/app.xml with application properties
        /// </summary>
        private void GenerateAppXml(PresentationDocument document, RelationshipIdManager relationshipIdManager)
        {
            // Note: AddExtendedFilePropertiesPart doesn't support custom relationship IDs
            var appPart = document.AddExtendedFilePropertiesPart();
            appPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties(
                new DocumentFormat.OpenXml.ExtendedProperties.TotalTime { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.Words { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.Application { Text = "Microsoft Office PowerPoint" },
                new DocumentFormat.OpenXml.ExtendedProperties.PresentationFormat { Text = "On-screen Show (4:3)" },
                new DocumentFormat.OpenXml.ExtendedProperties.Paragraphs { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.Slides { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.Notes { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.HiddenSlides { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.MultimediaClips { Text = "0" },
                new DocumentFormat.OpenXml.ExtendedProperties.ScaleCrop { Text = "false" },
                new DocumentFormat.OpenXml.ExtendedProperties.HeadingPairs(),
                new DocumentFormat.OpenXml.ExtendedProperties.TitlesOfParts(),
                new DocumentFormat.OpenXml.ExtendedProperties.Company { Text = "" },
                new DocumentFormat.OpenXml.ExtendedProperties.LinksUpToDate { Text = "false" },
                new DocumentFormat.OpenXml.ExtendedProperties.SharedDocument { Text = "false" },
                new DocumentFormat.OpenXml.ExtendedProperties.HyperlinksChanged { Text = "false" },
                new DocumentFormat.OpenXml.ExtendedProperties.ApplicationVersion { Text = "16.0000" }
            );
        }

        /// <summary>
        /// Generates docProps/core.xml with core document properties
        /// </summary>
        private void GenerateCoreXml(PresentationDocument document)
        {
            // PackageProperties are automatically created when accessing them
            var props = document.PackageProperties;
            props.Creator = "OfficeHelperOpenXml";
            props.LastModifiedBy = "OfficeHelperOpenXml";
            props.Created = DateTime.UtcNow;
            props.Modified = DateTime.UtcNow;
            props.Revision = "1";
        }

        /// <summary>
        /// Generates ppt/presProps.xml with presentation properties
        /// </summary>
        private void GeneratePresProps(PresentationDocument document, RelationshipIdManager relationshipIdManager)
        {
            var presentationPart = document.PresentationPart;
            if (presentationPart.PresentationPropertiesPart == null)
            {
                var presPropsPart = presentationPart.AddNewPart<PresentationPropertiesPart>(relationshipIdManager.GetNextId());
                presPropsPart.PresentationProperties = new PresentationProperties();
                
                // Apply namespace declarations using utility class
                Writers.NamespaceDeclarationApplier.ApplyPresPropsNamespaces(presPropsPart.PresentationProperties);
            }
        }

        /// <summary>
        /// Generates ppt/viewProps.xml with view properties
        /// </summary>
        private void GenerateViewProps(PresentationDocument document, RelationshipIdManager relationshipIdManager)
        {
            var presentationPart = document.PresentationPart;
            if (presentationPart.ViewPropertiesPart == null)
            {
                var viewPropsPart = presentationPart.AddNewPart<ViewPropertiesPart>(relationshipIdManager.GetNextId());
                viewPropsPart.ViewProperties = new ViewProperties(
                    new NormalViewProperties(
                        new RestoredLeft { Size = 15620 },
                        new RestoredTop { Size = 94660 }
                    ),
                    new SlideViewProperties(),
                    new NotesTextViewProperties(),
                    new GridSpacing { Cx = 72008, Cy = 72008 }
                );
            }
        }

        /// <summary>
        /// Generates ppt/tableStyles.xml with table style definitions
        /// </summary>
        private void GenerateTableStyles(PresentationDocument document, RelationshipIdManager relationshipIdManager)
        {
            var presentationPart = document.PresentationPart;
            if (presentationPart.TableStylesPart == null)
            {
                var tableStylesPart = presentationPart.AddNewPart<TableStylesPart>(relationshipIdManager.GetNextId());
                tableStylesPart.TableStyleList = new A.TableStyleList { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            }
        }
    }
}
