using System;
using System.IO;
using Xunit;
using OfficeHelperOpenXml.Core.Comparison;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Tests for ComparisonOrchestrator component
    /// </summary>
    public class ComparisonOrchestratorTests
    {
        [Fact]
        public void Constructor_WithNullExtractor_ThrowsArgumentNullException()
        {
            // Arrange
            var logger = new ConversionLogger();
            var tempDirManager = new TempDirManager(logger);
            var fileComparator = new FileComparator();
            var xmlComparator = new XmlComparator();
            var reportGenerator = new ReportGenerator(logger);

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => new ComparisonOrchestrator(
                null,
                fileComparator,
                xmlComparator,
                reportGenerator,
                tempDirManager,
                logger));
        }

        [Fact]
        public void Constructor_WithNullFileComparator_ThrowsArgumentNullException()
        {
            // Arrange
            var logger = new ConversionLogger();
            var tempDirManager = new TempDirManager(logger);
            var extractor = new PptxExtractor(tempDirManager, logger);
            var xmlComparator = new XmlComparator();
            var reportGenerator = new ReportGenerator(logger);

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => new ComparisonOrchestrator(
                extractor,
                null,
                xmlComparator,
                reportGenerator,
                tempDirManager,
                logger));
        }

        [Fact]
        public void Constructor_WithNullXmlComparator_ThrowsArgumentNullException()
        {
            // Arrange
            var logger = new ConversionLogger();
            var tempDirManager = new TempDirManager(logger);
            var extractor = new PptxExtractor(tempDirManager, logger);
            var fileComparator = new FileComparator();
            var reportGenerator = new ReportGenerator(logger);

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => new ComparisonOrchestrator(
                extractor,
                fileComparator,
                null,
                reportGenerator,
                tempDirManager,
                logger));
        }

        [Fact]
        public void Constructor_WithNullReportGenerator_ThrowsArgumentNullException()
        {
            // Arrange
            var logger = new ConversionLogger();
            var tempDirManager = new TempDirManager(logger);
            var extractor = new PptxExtractor(tempDirManager, logger);
            var fileComparator = new FileComparator();
            var xmlComparator = new XmlComparator();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => new ComparisonOrchestrator(
                extractor,
                fileComparator,
                xmlComparator,
                null,
                tempDirManager,
                logger));
        }

        [Fact]
        public void Constructor_WithNullTempDirManager_ThrowsArgumentNullException()
        {
            // Arrange
            var logger = new ConversionLogger();
            var tempDirManager = new TempDirManager(logger);
            var extractor = new PptxExtractor(tempDirManager, logger);
            var fileComparator = new FileComparator();
            var xmlComparator = new XmlComparator();
            var reportGenerator = new ReportGenerator(logger);

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => new ComparisonOrchestrator(
                extractor,
                fileComparator,
                xmlComparator,
                reportGenerator,
                null,
                logger));
        }

        [Fact]
        public void Constructor_WithValidParameters_CreatesInstance()
        {
            // Arrange
            var logger = new ConversionLogger();
            var tempDirManager = new TempDirManager(logger);
            var extractor = new PptxExtractor(tempDirManager, logger);
            var fileComparator = new FileComparator();
            var xmlComparator = new XmlComparator();
            var reportGenerator = new ReportGenerator(logger);

            // Act
            var orchestrator = new ComparisonOrchestrator(
                extractor,
                fileComparator,
                xmlComparator,
                reportGenerator,
                tempDirManager,
                logger);

            // Assert
            Assert.NotNull(orchestrator);
        }

        [Fact]
        public void Constructor_DefaultConstructor_CreatesInstance()
        {
            // Act
            var orchestrator = new ComparisonOrchestrator();

            // Assert
            Assert.NotNull(orchestrator);
        }

        [Fact]
        public void ExecuteComparison_WithNullConfig_ThrowsArgumentNullException()
        {
            // Arrange
            var orchestrator = new ComparisonOrchestrator();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => orchestrator.ExecuteComparison(null));
        }

        [Fact]
        public void ExecuteComparison_WithEmptyConfig_CreatesOutputDirectory()
        {
            // Arrange
            var orchestrator = new ComparisonOrchestrator();
            var outputDir = Path.Combine(Path.GetTempPath(), $"test_output_{Guid.NewGuid():N}");
            
            var config = new ComparisonConfig
            {
                OutputDirectory = outputDir
            };

            try
            {
                // Act
                var summary = orchestrator.ExecuteComparison(config);

                // Assert
                Assert.NotNull(summary);
                Assert.True(Directory.Exists(outputDir), "Output directory should be created");
            }
            finally
            {
                // Cleanup
                if (Directory.Exists(outputDir))
                {
                    Directory.Delete(outputDir, true);
                }
            }
        }
    }
}
