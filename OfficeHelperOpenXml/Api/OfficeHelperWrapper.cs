using System;
using System.IO;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// OfficeHelper包装类 - 提供简单的静态方法，便于从Python和C++调用
    /// </summary>
    public static class OfficeHelperWrapper
    {
        private static readonly string Version = "2.0.0-OpenXML";

        /// <summary>
        /// 分析PowerPoint文件并返回JSON字符串
        /// </summary>
        /// <param name="filePath">PowerPoint文件路径</param>
        /// <returns>JSON字符串</returns>
        public static string AnalyzePowerPoint(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath))
                {
                    return "{\"error\":\"文件路径不能为空\"}";
                }

                if (!File.Exists(filePath))
                {
                    return $"{{\"error\":\"文件不存在: {filePath}\"}}";
                }

                using (var reader = new PowerPointReader())
                {
                    if (!reader.Load(filePath))
                    {
                        return "{\"error\":\"加载文件失败\"}";
                    }
                    return reader.ToJson();
                }
            }
            catch (Exception ex)
            {
                return $"{{\"error\":\"分析文件时出错: {ex.Message}\"}}";
            }
        }

        /// <summary>
        /// 分析PowerPoint文件并保存为JSON文件
        /// </summary>
        /// <param name="filePath">PowerPoint文件路径</param>
        /// <param name="outputPath">输出JSON文件路径</param>
        /// <returns>是否成功</returns>
        public static bool AnalyzePowerPointToFile(string filePath, string outputPath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(outputPath))
                {
                    return false;
                }

                if (!File.Exists(filePath))
                {
                    return false;
                }

                using (var reader = new PowerPointReader())
                {
                    if (!reader.Load(filePath))
                    {
                        return false;
                    }
                    return reader.SaveToJson(outputPath);
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 检查文件是否存在
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>文件是否存在</returns>
        public static bool FileExists(string filePath)
        {
            return File.Exists(filePath);
        }

        /// <summary>
        /// 获取版本信息
        /// </summary>
        /// <returns>版本字符串</returns>
        public static string GetVersion()
        {
            return Version;
        }

        /// <summary>
        /// 获取库信息
        /// </summary>
        /// <returns>库信息JSON</returns>
        public static string GetLibraryInfo()
        {
            return $"{{\"name\":\"OfficeHelperOpenXml\",\"version\":\"{Version}\",\"framework\":\"OpenXML SDK\",\"targetFramework\":\"netstandard2.0\"}}";
        }
    }
}
