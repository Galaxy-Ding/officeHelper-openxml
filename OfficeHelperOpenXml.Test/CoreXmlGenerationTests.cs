using System;
using System.IO;
using System.Linq;
using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for core.xml generation in PackageMetadataGenerator
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
    /// **Property 4: Core.xml exists with required properties**
    /// </summary>
    public class CoreXmlGenerationTests : IDisposable
    {
        private readonly string _testOutputPath;

        public CoreXmlGenerationTests()
        {
            _testOutputPath = Path.Combine(Path.GetTempPath(), $"test_core_{Guid.NewGuid()}.pptx");
        }

        public void Dispose()
        {
            if (File.Exists(_testOutputPath))
            {
                try
                {
                    File.Delete(_testOutputPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        [Fact]
        public void GenerateCoreXml_CreatesPackageProperties()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert
                Assert.NotNull(document.PackageProperties);
            }
        }

        [Fact]
        public void GenerateCoreXml_ContainsCreatorProperty()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert
                Assert.NotNull(document.PackageProperties.Creator);
                Assert.Equal("OfficeHelperOpenXml", document.PackageProperties.Creator);
            }
        }

        [Fact]
        public void GenerateCoreXml_ContainsLastModifiedByProperty()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert
                Assert.NotNull(document.PackageProperties.LastModifiedBy);
                Assert.Equal("OfficeHelperOpenXml", document.PackageProperties.LastModifiedBy);
            }
        }

        [Fact]
        public void GenerateCoreXml_ContainsCreatedProperty()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert
                Assert.NotNull(document.PackageProperties.Created);
                // Verify it's a recent timestamp (within last minute)
                var timeDiff = DateTime.UtcNow - document.PackageProperties.Created.Value;
                Assert.True(timeDiff.TotalMinutes < 1, "Created timestamp should be recent");
            }
        }

        [Fact]
        public void GenerateCoreXml_ContainsModifiedProperty()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert
                Assert.NotNull(document.PackageProperties.Modified);
                // Verify it's a recent timestamp (within last minute)
                var timeDiff = DateTime.UtcNow - document.PackageProperties.Modified.Value;
                Assert.True(timeDiff.TotalMinutes < 1, "Modified timestamp should be recent");
            }
        }

        [Fact]
        public void GenerateCoreXml_ContainsRevisionProperty()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert
                Assert.NotNull(document.PackageProperties.Revision);
                Assert.Equal("1", document.PackageProperties.Revision);
            }
        }

        [Fact]
        public void GenerateCoreXml_AllRequiredPropertiesPresent()
        {
            // Arrange
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();

                // Act
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);

                // Assert - Verify all required properties are present
                var props = document.PackageProperties;
                Assert.NotNull(props.Creator);
                Assert.NotNull(props.LastModifiedBy);
                Assert.NotNull(props.Created);
                Assert.NotNull(props.Modified);
                Assert.NotNull(props.Revision);
            }
        }

        [Fact]
        public void GenerateCoreXml_CoreFileExistsInPackage()
        {
            // Arrange & Act
            using (var document = PresentationDocument.Create(_testOutputPath, PresentationDocumentType.Presentation))
            {
                document.AddPresentationPart();
                var metadataGenerator = new PackageMetadataGenerator();
                var relationshipIdManager = new RelationshipIdManager();
                metadataGenerator.GenerateAllMetadata(document, relationshipIdManager);
                document.Save();
            }

            // Assert - Reopen and verify core.xml exists
            using (var document = PresentationDocument.Open(_testOutputPath, false))
            {
                Assert.NotNull(document.PackageProperties);
                Assert.NotNull(document.PackageProperties.Creator);
            }
        }
    }
}
