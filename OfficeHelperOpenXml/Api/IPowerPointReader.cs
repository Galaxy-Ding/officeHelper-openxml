using System;
using System.Collections.Generic;
using OfficeHelperOpenXml.Core.Readers;
using OfficeHelperOpenXml.Interfaces;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// PowerPoint读取器接口 - 公共API
    /// </summary>
    public interface IPowerPointReader : IDisposable
    {
        /// <summary>
        /// 文件路径
        /// </summary>
        string FilePath { get; }
        
        /// <summary>
        /// 是否已加载
        /// </summary>
        bool IsLoaded { get; }
        
        /// <summary>
        /// 演示文稿信息
        /// </summary>
        PresentationInfo PresentationInfo { get; }
        
        /// <summary>
        /// 加载PowerPoint文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>是否加载成功</returns>
        bool Load(string filePath);
        
        /// <summary>
        /// 重新加载当前文件
        /// </summary>
        /// <returns>是否重新加载成功</returns>
        bool Reload();
        
        /// <summary>
        /// 获取JSON格式的分析结果
        /// </summary>
        /// <returns>JSON字符串</returns>
        string ToJson();
        
        /// <summary>
        /// 保存分析结果为JSON文件
        /// </summary>
        /// <param name="outputPath">输出文件路径</param>
        /// <returns>是否保存成功</returns>
        bool SaveToJson(string outputPath);
        
        /// <summary>
        /// 获取指定页面的元素
        /// </summary>
        /// <param name="pageNumber">页面编号（从1开始）</param>
        /// <returns>页面元素列表</returns>
        List<IElement> GetPageElements(int pageNumber);
        
        /// <summary>
        /// 获取所有页面元素
        /// </summary>
        /// <returns>所有页面元素</returns>
        List<IElement> GetAllElements();
    }
}
