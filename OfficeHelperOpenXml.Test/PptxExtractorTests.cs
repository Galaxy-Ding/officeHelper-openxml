using System;
using System.IO;
using System.Linq;
using OfficeHelperOpenXml.Core.Comparison;
using Xunit;

namespace OfficeHelperOpenXml.Test
{
    /// <summary>
    /// Unit tests for PptxExtractor component
    /// </summary>
    public class PptxExtractorTests : IDisposable
    {
        private readonly TempDirManager _tempDirManager;
        private readonly PptxExtractor _extractor;

        public PptxExtractorTests()
        {
            _tempDirManager = new TempDirManager();
            _extractor = new PptxExtractor(_tempDirManager);
        }

        public void Dispose()
        {
            // Cleanup temporary directories after each test
            _tempDirManager.CleanupAll();
        }

        [Fact]
        public void Extract_WithValidPptx_ShouldSucceed()
        {
            // Arrange
            string testFile = @"test_ppt\textbox_1_content.pptx";
            
            // Skip test if file doesn't exist
            if (!File.Exists(testFile))
            {
                return;
            }

            // Act
            var result = _extractor.Extract(testFile);

            // Assert
            Assert.True(result.Success, $"Extraction failed: {result.ErrorMessage}");
            Assert.NotNull(result.ExtractedPath);
            Assert.True(Directory.Exists(result.ExtractedPath));
            Assert.NotEmpty(result.Files);
            Assert.Null(result.ErrorMessage);
        }

        [Fact]
        public void Extract_WithNullPath_ShouldReturnFailure()
        {
            // Act
            var result = _extractor.Extract(null);

            // Assert
            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Contains("cannot be null or empty", result.ErrorMessage);
        }

        [Fact]
        public void Extract_WithEmptyPath_ShouldReturnFailure()
        {
            // Act
            var result = _extractor.Extract("");

            // Assert
            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Contains("cannot be null or empty", result.ErrorMessage);
        }

        [Fact]
        public void Extract_WithNonExistentFile_ShouldReturnFailure()
        {
            // Arrange
            string nonExistentFile = "nonexistent_file.pptx";

            // Act
            var result = _extractor.Extract(nonExistentFile);

            // Assert
            Assert.False(result.Success);
            Assert.NotNull(result.ErrorMessage);
            Assert.Contains("not found", result.ErrorMessage);
        }

        [Fact]
        public void Extract_ShouldPreserveDirectoryStructure()
        {
            // Arrange
            string testFile = @"test_ppt\textbox_1_content.pptx";
            
            // Skip test if file doesn't exist
            if (!File.Exists(testFile))
            {
                return;
            }

            // Act
            var result = _extractor.Extract(testFile);

            // Assert
            Assert.True(result.Success);
            
            // PPTX files should contain standard directories
            var hasContentTypes = result.Files.Any(f => f.Contains("[Content_Types].xml"));
            var hasRels = result.Files.Any(f => f.Contains("_rels"));
            
            Assert.True(hasContentTypes || hasRels, "Should preserve standard PPTX structure");
        }

        [Fact]
        public void GetFileList_WithValidDirectory_ShouldReturnAllFiles()
        {
            // Arrange
            string testFile = @"test_ppt\textbox_1_content.pptx";
            
            // Skip test if file doesn't exist
            if (!File.Exists(testFile))
            {
                return;
            }

            var extractResult = _extractor.Extract(testFile);
            Assert.True(extractResult.Success);

            // Act
            var fileList = _extractor.GetFileList(extractResult.ExtractedPath);

            // Assert
            Assert.NotEmpty(fileList);
            Assert.All(fileList, file => Assert.False(Path.IsPathRooted(file), "Paths should be relative"));
        }

        [Fact]
        public void GetFileList_WithNullPath_ShouldThrowArgumentException()
        {
            // Act & Assert
            Assert.Throws<ArgumentException>(() => _extractor.GetFileList(null));
        }

        [Fact]
        public void GetFileList_WithNonExistentDirectory_ShouldThrowDirectoryNotFoundException()
        {
            // Arrange
            string nonExistentDir = "nonexistent_directory";

            // Act & Assert
            Assert.Throws<DirectoryNotFoundException>(() => _extractor.GetFileList(nonExistentDir));
        }

        [Fact]
        public void GetXmlFiles_ShouldFilterXmlAndRelsFiles()
        {
            // Arrange
            string testFile = @"test_ppt\textbox_1_content.pptx";
            
            // Skip test if file doesn't exist
            if (!File.Exists(testFile))
            {
                return;
            }

            var extractResult = _extractor.Extract(testFile);
            Assert.True(extractResult.Success);

            // Act
            var xmlFiles = _extractor.GetXmlFiles(extractResult.ExtractedPath);

            // Assert
            Assert.NotEmpty(xmlFiles);
            Assert.All(xmlFiles, file => 
            {
                Assert.True(
                    file.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                    file.EndsWith(".rels", StringComparison.OrdinalIgnoreCase),
                    $"File {file} should be XML or RELS");
            });
        }

        [Fact]
        public void Extract_WithWindowsAbsolutePath_ShouldHandleCorrectly()
        {
            // Arrange
            string testFile = Path.GetFullPath(@"test_ppt\textbox_1_content.pptx");
            
            // Skip test if file doesn't exist
            if (!File.Exists(testFile))
            {
                return;
            }

            // Act
            var result = _extractor.Extract(testFile);

            // Assert
            Assert.True(result.Success, $"Should handle Windows absolute path: {result.ErrorMessage}");
            Assert.NotEmpty(result.Files);
        }

        [Fact]
        public void Extract_MultipleTimes_ShouldUseSeparateDirectories()
        {
            // Arrange
            string testFile = @"test_ppt\textbox_1_content.pptx";
            
            // Skip test if file doesn't exist
            if (!File.Exists(testFile))
            {
                return;
            }

            // Act
            var result1 = _extractor.Extract(testFile);
            var result2 = _extractor.Extract(testFile);

            // Assert
            Assert.True(result1.Success);
            Assert.True(result2.Success);
            Assert.NotEqual(result1.ExtractedPath, result2.ExtractedPath);
            Assert.Equal(2, _tempDirManager.TrackedDirectoryCount);
        }
    }
}
