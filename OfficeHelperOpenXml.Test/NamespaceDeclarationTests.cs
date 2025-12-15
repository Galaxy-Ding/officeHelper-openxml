using System;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Core.Writers;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests for NamespaceDeclarationApplier utility class.
    /// Validates: Requirements 7.1, 7.2, 7.3, 7.4, 7.5
    /// Property 6: All XML files have namespace declarations
    /// </summary>
    public class NamespaceDeclarationTests
    {
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";

        [Fact]
        public void ApplyPresentationNamespaces_AddsAllRequiredNamespaces()
        {
            // Arrange
            var presentation = new Presentation();

            // Act
            NamespaceDeclarationApplier.ApplyPresentationNamespaces(presentation);

            // Assert
            var namespaces = presentation.NamespaceDeclarations.ToList();
            Assert.Contains(namespaces, ns => ns.Key == "a" && ns.Value == DrawingNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "r" && ns.Value == RelationshipsNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "p" && ns.Value == PresentationNamespace);
        }

        [Fact]
        public void ApplySlideNamespaces_AddsAllRequiredNamespaces()
        {
            // Arrange
            var slide = new Slide();

            // Act
            NamespaceDeclarationApplier.ApplySlideNamespaces(slide);

            // Assert
            var namespaces = slide.NamespaceDeclarations.ToList();
            Assert.Contains(namespaces, ns => ns.Key == "a" && ns.Value == DrawingNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "r" && ns.Value == RelationshipsNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "p" && ns.Value == PresentationNamespace);
        }

        [Fact]
        public void ApplySlideMasterNamespaces_AddsAllRequiredNamespaces()
        {
            // Arrange
            var slideMaster = new SlideMaster();

            // Act
            NamespaceDeclarationApplier.ApplySlideMasterNamespaces(slideMaster);

            // Assert
            var namespaces = slideMaster.NamespaceDeclarations.ToList();
            Assert.Contains(namespaces, ns => ns.Key == "a" && ns.Value == DrawingNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "r" && ns.Value == RelationshipsNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "p" && ns.Value == PresentationNamespace);
        }

        [Fact]
        public void ApplySlideLayoutNamespaces_AddsAllRequiredNamespaces()
        {
            // Arrange
            var slideLayout = new SlideLayout();

            // Act
            NamespaceDeclarationApplier.ApplySlideLayoutNamespaces(slideLayout);

            // Assert
            var namespaces = slideLayout.NamespaceDeclarations.ToList();
            Assert.Contains(namespaces, ns => ns.Key == "a" && ns.Value == DrawingNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "r" && ns.Value == RelationshipsNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "p" && ns.Value == PresentationNamespace);
        }

        [Fact]
        public void ApplyPresPropsNamespaces_AddsAllRequiredNamespaces()
        {
            // Arrange
            var presProps = new PresentationProperties();

            // Act
            NamespaceDeclarationApplier.ApplyPresPropsNamespaces(presProps);

            // Assert
            var namespaces = presProps.NamespaceDeclarations.ToList();
            Assert.Contains(namespaces, ns => ns.Key == "a" && ns.Value == DrawingNamespace);
            Assert.Contains(namespaces, ns => ns.Key == "r" && ns.Value == RelationshipsNamespace);
        }

        [Fact]
        public void ApplyPresentationNamespaces_DoesNotDuplicateExistingNamespaces()
        {
            // Arrange
            var presentation = new Presentation();
            presentation.AddNamespaceDeclaration("a", DrawingNamespace);

            // Act
            NamespaceDeclarationApplier.ApplyPresentationNamespaces(presentation);

            // Assert
            var namespaces = presentation.NamespaceDeclarations.ToList();
            var aNamespaces = namespaces.Where(ns => ns.Key == "a").ToList();
            Assert.Single(aNamespaces); // Should only have one "a" namespace
        }

        [Fact]
        public void ApplySlideNamespaces_DoesNotDuplicateExistingNamespaces()
        {
            // Arrange
            var slide = new Slide();
            slide.AddNamespaceDeclaration("p", PresentationNamespace);

            // Act
            NamespaceDeclarationApplier.ApplySlideNamespaces(slide);

            // Assert
            var namespaces = slide.NamespaceDeclarations.ToList();
            var pNamespaces = namespaces.Where(ns => ns.Key == "p").ToList();
            Assert.Single(pNamespaces); // Should only have one "p" namespace
        }

        [Fact]
        public void ApplyPresentationNamespaces_HandlesNullGracefully()
        {
            // Act & Assert - Should not throw
            NamespaceDeclarationApplier.ApplyPresentationNamespaces(null);
        }

        [Fact]
        public void ApplySlideNamespaces_HandlesNullGracefully()
        {
            // Act & Assert - Should not throw
            NamespaceDeclarationApplier.ApplySlideNamespaces(null);
        }

        [Fact]
        public void ApplySlideMasterNamespaces_HandlesNullGracefully()
        {
            // Act & Assert - Should not throw
            NamespaceDeclarationApplier.ApplySlideMasterNamespaces(null);
        }

        [Fact]
        public void ApplySlideLayoutNamespaces_HandlesNullGracefully()
        {
            // Act & Assert - Should not throw
            NamespaceDeclarationApplier.ApplySlideLayoutNamespaces(null);
        }

        [Fact]
        public void ApplyPresPropsNamespaces_HandlesNullGracefully()
        {
            // Act & Assert - Should not throw
            NamespaceDeclarationApplier.ApplyPresPropsNamespaces(null);
        }
    }
}
