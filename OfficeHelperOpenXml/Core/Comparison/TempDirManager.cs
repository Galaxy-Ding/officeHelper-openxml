using System;
using System.Collections.Generic;
using System.IO;
using OfficeHelperOpenXml.Core.Converters;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Manages temporary directories for PPTX extraction and comparison
    /// </summary>
    public class TempDirManager
    {
        private readonly List<string> _tempDirectories;
        private readonly ConversionLogger _logger;

        public TempDirManager()
        {
            _tempDirectories = new List<string>();
            _logger = new ConversionLogger();
        }

        public TempDirManager(ConversionLogger logger)
        {
            _tempDirectories = new List<string>();
            _logger = logger ?? new ConversionLogger();
        }

        /// <summary>
        /// Creates a new temporary directory with the specified prefix
        /// </summary>
        /// <param name="prefix">Prefix for the temporary directory name</param>
        /// <returns>Full path to the created temporary directory</returns>
        public string CreateTempDirectory(string prefix)
        {
            if (string.IsNullOrWhiteSpace(prefix))
            {
                prefix = "temp";
            }

            // Create a unique directory name using prefix and GUID
            string tempPath = Path.Combine(Path.GetTempPath(), $"{prefix}_{Guid.NewGuid():N}");

            try
            {
                Directory.CreateDirectory(tempPath);
                TrackDirectory(tempPath);
                _logger.LogInfo($"Created temporary directory: {tempPath}");
                return tempPath;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to create temporary directory: {ex.Message}");
                throw new ComparisonException($"Failed to create temporary directory with prefix '{prefix}'", ex)
                {
                    Context = "TempDirManager.CreateTempDirectory",
                    FilePath = tempPath,
                    Category = ErrorCategory.FileSystem
                };
            }
        }

        /// <summary>
        /// Tracks an existing directory for cleanup
        /// </summary>
        /// <param name="path">Path to the directory to track</param>
        public void TrackDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Directory path cannot be null or empty", nameof(path));
            }

            if (!_tempDirectories.Contains(path))
            {
                _tempDirectories.Add(path);
            }
        }

        /// <summary>
        /// Cleans up all tracked temporary directories
        /// </summary>
        /// <returns>Result of the cleanup operation</returns>
        public CleanupResult CleanupAll()
        {
            var result = new CleanupResult
            {
                TotalDirectories = _tempDirectories.Count
            };

            _logger.LogInfo($"Starting cleanup of {result.TotalDirectories} temporary directories");

            foreach (var directory in _tempDirectories.ToArray())
            {
                if (CleanupDirectory(directory))
                {
                    result.SuccessfulCleanups++;
                }
                else
                {
                    result.FailedCleanups++;
                    result.Errors.Add($"Failed to cleanup directory: {directory}");
                }
            }

            if (result.FailedCleanups > 0)
            {
                _logger.LogWarning($"Cleanup completed with {result.FailedCleanups} failures out of {result.TotalDirectories} directories");
            }
            else
            {
                _logger.LogSuccess($"Successfully cleaned up all {result.SuccessfulCleanups} temporary directories");
            }

            return result;
        }

        /// <summary>
        /// Cleans up a specific directory
        /// </summary>
        /// <param name="path">Path to the directory to cleanup</param>
        /// <returns>True if cleanup was successful, false otherwise</returns>
        public bool CleanupDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, recursive: true);
                    _logger.LogInfo($"Deleted temporary directory: {path}");
                }

                // Remove from tracking list
                _tempDirectories.Remove(path);
                return true;
            }
            catch (Exception ex)
            {
                // Log the error but don't throw - cleanup errors should not fail the overall process
                _logger.LogWarning($"Failed to delete directory '{path}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets the count of currently tracked directories
        /// </summary>
        public int TrackedDirectoryCount => _tempDirectories.Count;

        /// <summary>
        /// Gets a read-only copy of tracked directories
        /// </summary>
        public IReadOnlyList<string> TrackedDirectories => _tempDirectories.AsReadOnly();
    }
}
