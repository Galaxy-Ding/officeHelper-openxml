using System;
using System.IO;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// PowerPoint读取器工厂类
    /// </summary>
    public static class PowerPointReaderFactory
    {
        /// <summary>
        /// 创建并加载PowerPoint读取器
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="success">是否成功</param>
        /// <returns>PowerPointReader实例</returns>
        public static IPowerPointReader CreateReader(string filePath, out bool success)
        {
            var reader = new PowerPointReader();
            success = reader.Load(filePath);
            return reader;
        }

        /// <summary>
        /// 创建PowerPoint读取器（不加载文件）
        /// </summary>
        /// <returns>PowerPointReader实例</returns>
        public static IPowerPointReader CreateReader()
        {
            return new PowerPointReader();
        }

        /// <summary>
        /// 快速读取并返回JSON
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>JSON字符串</returns>
        public static string QuickReadToJson(string filePath)
        {
            using (var reader = new PowerPointReader())
            {
                return reader.ReadToJson(filePath);
            }
        }
    }
}
