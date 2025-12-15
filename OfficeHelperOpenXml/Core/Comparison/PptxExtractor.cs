using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Extracts PPTX files and provides file listing functionality
    /// </summary>
    public class PptxExtractor
    {
        private readonly TempDirManager _tempDirManager;
        private readonly ConversionLogger _logger;

        public PptxExtractor(TempDirManager tempDirManager)
        {
            _tempDirManager = tempDirManager ?? throw new ArgumentNullException(nameof(tempDirManager));
            _logger = new ConversionLogger();
        }

        public PptxExtractor(TempDirManager tempDirManager, ConversionLogger logger)
        {
            _tempDirManager = tempDirManager ?? throw new ArgumentNullException(nameof(tempDirManager));
            _logger = logger ?? new ConversionLogger();
        }

        /// <summary>
        /// Extracts a PPTX file to a temporary directory
        /// </summary>
        /// <param name="pptxPath">Path to the PPTX file</param>
        /// <returns>ExtractionResult containing success status, extracted path, and file list</returns>
        public ExtractionResult Extract(string pptxPath)
        {
            var result = new ExtractionResult();

            // Validate input path
            if (string.IsNullOrWhiteSpace(pptxPath))
            {
                result.Success = false;
                result.ErrorMessage = "PPTX file path cannot be null or empty";
                _logger.LogError(result.ErrorMessage);
                return result;
            }

            // Check if file exists
            if (!File.Exists(pptxPath))
            {
                result.Success = false;
                result.ErrorMessage = $"PPTX file not found: {pptxPath}";
                _logger.LogError(result.ErrorMessage);
                return result;
            }

            try
            {
                // Create temporary directory for extraction
                string fileName = Path.GetFileNameWithoutExtension(pptxPath);
                string tempDir = _tempDirManager.CreateTempDirectory($"pptx_extract_{fileName}");

                _logger.LogInfo($"Extracting PPTX file: {pptxPath}");
                _logger.LogInfo($"Extraction directory: {tempDir}");

                // Extract the PPTX (ZIP) file
                try
                {
                    ZipFile.ExtractToDirectory(pptxPath, tempDir);
                }
                catch (InvalidDataException ex)
                {
                    // Corrupt ZIP file
                    result.Success = false;
                    result.ErrorMessage = $"Corrupt or invalid PPTX file: {pptxPath}. Error: {ex.Message}";
                    _logger.LogError(result.ErrorMessage);
                    
                    throw new ComparisonException(result.ErrorMessage, ex)
                    {
                        Context = "PptxExtractor.Extract",
                        FilePath = pptxPath,
                        Category = ErrorCategory.Archive
                    };
                }

                // Get list of all extracted files
                result.Files = GetFileList(tempDir);
                result.ExtractedPath = tempDir;
                result.Success = true;

                _logger.LogSuccess($"Successfully extracted {result.Files.Count} files from PPTX");

                return result;
            }
            catch (ComparisonException)
            {
                // Re-throw comparison exceptions
                throw;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to extract PPTX file '{pptxPath}': {ex.Message}";
                _logger.LogError(result.ErrorMessage);

                throw new ComparisonException(result.ErrorMessage, ex)
                {
                    Context = "PptxExtractor.Extract",
                    FilePath = pptxPath,
                    Category = ErrorCategory.FileSystem
                };
            }
        }

        /// <summary>
        /// Recursively lists all files in a directory with relative paths
        /// </summary>
        /// <param name="extractedDir">Root directory to list files from</param>
        /// <returns>List of relative file paths</returns>
        public List<string> GetFileList(string extractedDir)
        {
            if (string.IsNullOrWhiteSpace(extractedDir))
            {
                throw new ArgumentException("Extracted directory path cannot be null or empty", nameof(extractedDir));
            }

            if (!Directory.Exists(extractedDir))
            {
                throw new DirectoryNotFoundException($"Directory not found: {extractedDir}");
            }

            var fileList = new List<string>();

            try
            {
                // Get all files recursively
                var allFiles = Directory.GetFiles(extractedDir, "*", SearchOption.AllDirectories);

                // Convert to relative paths
                foreach (var filePath in allFiles)
                {
                    string relativePath = GetRelativePath(extractedDir, filePath);
                    fileList.Add(relativePath);
                }

                // Sort for consistent ordering
                fileList.Sort(StringComparer.OrdinalIgnoreCase);

                return fileList;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to list files in directory '{extractedDir}': {ex.Message}");
                throw new ComparisonException($"Failed to list files in directory", ex)
                {
                    Context = "PptxExtractor.GetFileList",
                    FilePath = extractedDir,
                    Category = ErrorCategory.FileSystem
                };
            }
        }

        /// <summary>
        /// Filters and returns only XML and .rels files from a directory
        /// </summary>
        /// <param name="extractedDir">Root directory to search</param>
        /// <returns>List of relative paths to XML and .rels files</returns>
        public List<string> GetXmlFiles(string extractedDir)
        {
            var allFiles = GetFileList(extractedDir);

            // Filter for XML and .rels files
            var xmlFiles = allFiles
                .Where(f => f.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                           f.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                .ToList();

            _logger.LogInfo($"Found {xmlFiles.Count} XML/RELS files out of {allFiles.Count} total files");

            return xmlFiles;
        }

        /// <summary>
        /// Gets the relative path from a base directory to a file
        /// </summary>
        private string GetRelativePath(string basePath, string fullPath)
        {
            // Normalize paths
            basePath = Path.GetFullPath(basePath);
            fullPath = Path.GetFullPath(fullPath);

            // Ensure base path ends with directory separator
            if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                basePath += Path.DirectorySeparatorChar;
            }

            // Get relative path
            Uri baseUri = new Uri(basePath);
            Uri fullUri = new Uri(fullPath);
            Uri relativeUri = baseUri.MakeRelativeUri(fullUri);

            // Convert to string and normalize separators
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());
            relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);

            return relativePath;
        }
    }
}
