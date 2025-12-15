using System;
using System.Collections.Generic;
using OfficeHelperOpenXml.Interfaces;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// PowerPoint写入器接口 - 公共API
    /// </summary>
    public interface IPowerPointWriter : IDisposable
    {
        /// <summary>
        /// 文件路径
        /// </summary>
        string FilePath { get; }
        
        /// <summary>
        /// 是否已打开
        /// </summary>
        bool IsOpen { get; }
        
        /// <summary>
        /// 打开或创建PowerPoint文件
        /// </summary>
        bool OpenOrCreate(string filePath);
        
        /// <summary>
        /// 创建新演示文稿
        /// </summary>
        bool CreateNew();
        
        /// <summary>
        /// 保存文件
        /// </summary>
        bool Save();
        
        /// <summary>
        /// 另存为
        /// </summary>
        bool SaveAs(string filePath);
        
        /// <summary>
        /// 添加幻灯片
        /// </summary>
        int AddSlide();
        
        /// <summary>
        /// 删除指定的幻灯片
        /// </summary>
        bool DeleteSlide(int slideIndex);
        
        /// <summary>
        /// 在指定位置添加元素
        /// </summary>
        bool AddElement(int slideIndex, IElement element);
        
        /// <summary>
        /// 在指定位置添加多个元素
        /// </summary>
        bool AddElements(int slideIndex, IEnumerable<IElement> elements);
        
        /// <summary>
        /// 更新指定位置的元素
        /// </summary>
        bool UpdateElement(int slideIndex, int elementIndex, IElement element);
        
        /// <summary>
        /// 删除指定位置的元素
        /// </summary>
        bool DeleteElement(int slideIndex, int elementIndex);
        
        /// <summary>
        /// 获取幻灯片数量
        /// </summary>
        int GetSlideCount();
    }
}
