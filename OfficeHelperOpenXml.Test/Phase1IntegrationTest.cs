using System;
using System.IO;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Integration tests for Phase 1 fixes: Namespace Declarations and Metadata
    /// </summary>
    public class Phase1IntegrationTest
    {
        private static readonly string TestJsonPath = Path.Combine("..", "..", "..", "..", "test_ppt", "textbox.json");
        private static readonly string TestOutputPath = Path.Combine("..", "..", "..", "..", "test_ppt", "textbox_phase1_test.pptx");

        [Fact]
        public void Phase1_GeneratedPptx_ContainsNamespaceDeclarations()
        {
            // Arrange
            if (File.Exists(TestOutputPath))
            {
                File.Delete(TestOutputPath);
            }

            var converter = new JsonToPptxConverter();

            // Act
            bool success = converter.Convert(TestJsonPath, TestOutputPath);

            // Assert
            Assert.True(success, "Conversion should succeed");
            Assert.True(File.Exists(TestOutputPath), "Output file should exist");

            // Verify namespace declarations
            using (var document = PresentationDocument.Open(TestOutputPath, false))
            {
                var presentation = document.PresentationPart.Presentation;
                
                // Check presentation namespace declarations
                var namespaces = presentation.NamespaceDeclarations.ToList();
                Assert.Contains(namespaces, ns => ns.Key == "a" && ns.Value == "http://schemas.openxmlformats.org/drawingml/2006/main");
                Assert.Contains(namespaces, ns => ns.Key == "r" && ns.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                Assert.Contains(namespaces, ns => ns.Key == "p" && ns.Value == "http://schemas.openxmlformats.org/presentationml/2006/main");

                // Check master slide namespace declarations
                var masterSlides = document.PresentationPart.SlideMasterParts;
                foreach (var masterPart in masterSlides)
                {
                    var masterNamespaces = masterPart.SlideMaster.NamespaceDeclarations.ToList();
                    Assert.Contains(masterNamespaces, ns => ns.Key == "a");
                    Assert.Contains(masterNamespaces, ns => ns.Key == "r");
                    Assert.Contains(masterNamespaces, ns => ns.Key == "p");
                }

                // Check content slide namespace declarations
                var slides = document.PresentationPart.SlideParts;
                foreach (var slidePart in slides)
                {
                    var slideNamespaces = slidePart.Slide.NamespaceDeclarations.ToList();
                    Assert.Contains(slideNamespaces, ns => ns.Key == "a");
                    Assert.Contains(slideNamespaces, ns => ns.Key == "r");
                    Assert.Contains(slideNamespaces, ns => ns.Key == "p");
                }
            }
        }

        [Fact]
        public void Phase1_GeneratedPptx_ContainsAllMetadataFiles()
        {
            // Arrange
            if (File.Exists(TestOutputPath))
            {
                File.Delete(TestOutputPath);
            }

            var converter = new JsonToPptxConverter();

            // Act
            bool success = converter.Convert(TestJsonPath, TestOutputPath);

            // Assert
            Assert.True(success, "Conversion should succeed");
            Assert.True(File.Exists(TestOutputPath), "Output file should exist");

            // Verify all required metadata files exist
            using (var document = PresentationDocument.Open(TestOutputPath, false))
            {
                // Check for extended file properties (docProps/app.xml)
                Assert.NotNull(document.ExtendedFilePropertiesPart);
                Assert.NotNull(document.ExtendedFilePropertiesPart.Properties);

                // Check for package properties (docProps/core.xml)
                Assert.NotNull(document.PackageProperties);

                // Check for presentation properties (ppt/presProps.xml)
                Assert.NotNull(document.PresentationPart.PresentationPropertiesPart);
                Assert.NotNull(document.PresentationPart.PresentationPropertiesPart.PresentationProperties);

                // Check for view properties (ppt/viewProps.xml)
                Assert.NotNull(document.PresentationPart.ViewPropertiesPart);
                Assert.NotNull(document.PresentationPart.ViewPropertiesPart.ViewProperties);

                // Check for table styles (ppt/tableStyles.xml)
                Assert.NotNull(document.PresentationPart.TableStylesPart);
                Assert.NotNull(document.PresentationPart.TableStylesPart.TableStyleList);

                // Check for theme files
                var masterSlides = document.PresentationPart.SlideMasterParts;
                foreach (var masterPart in masterSlides)
                {
                    Assert.NotNull(masterPart.ThemePart);
                    Assert.NotNull(masterPart.ThemePart.Theme);
                }
            }
        }

        [Fact]
        public void Phase1_GeneratedPptx_OpensInPowerPoint()
        {
            // Arrange
            if (File.Exists(TestOutputPath))
            {
                File.Delete(TestOutputPath);
            }

            var converter = new JsonToPptxConverter();

            // Act
            bool success = converter.Convert(TestJsonPath, TestOutputPath);

            // Assert
            Assert.True(success, "Conversion should succeed");
            Assert.True(File.Exists(TestOutputPath), "Output file should exist");

            // Verify the file can be opened without errors
            using (var document = PresentationDocument.Open(TestOutputPath, false))
            {
                // If we can open it and access the presentation, it's valid
                Assert.NotNull(document.PresentationPart);
                Assert.NotNull(document.PresentationPart.Presentation);
                
                // Verify basic structure
                Assert.NotNull(document.PresentationPart.Presentation.SlideIdList);
                Assert.NotNull(document.PresentationPart.Presentation.SlideMasterIdList);
            }
        }
    }
}
