using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using OfficeHelperOpenXml.Core.Readers;
using OfficeHelperOpenXml.Core.Converters;
using OfficeHelperOpenXml.Interfaces;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// PowerPoint读取器 - 基于OpenXML实现
    /// </summary>
    public class PowerPointReader : IPowerPointReader
    {
        private readonly PresentationReader _reader;
        private PresentationInfo _info;
        private string _filePath;
        private bool _disposed;

        public PowerPointReader()
        {
            _reader = new PresentationReader();
        }

        public string FilePath => _filePath;
        public bool IsLoaded => _info != null && string.IsNullOrEmpty(_info.Error);
        public PresentationInfo PresentationInfo => _info;

        public bool Load(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            if (!File.Exists(filePath))
                return false;

            try
            {
                _filePath = filePath;
                _info = _reader.ReadPresentation(filePath);
                return string.IsNullOrEmpty(_info.Error);
            }
            catch
            {
                return false;
            }
        }

        public bool Reload()
        {
            if (string.IsNullOrEmpty(_filePath))
                return false;
            return Load(_filePath);
        }

        public string ToJson()
        {
            if (_info == null)
                return "{\"error\":\"未加载文件\"}";
            return JsonConverter.ConvertToJson(_info);
        }

        public bool SaveToJson(string outputPath)
        {
            if (_info == null || string.IsNullOrEmpty(outputPath))
                return false;

            try
            {
                var json = JsonConverter.ConvertToJson(_info);
                File.WriteAllText(outputPath, json);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public List<IElement> GetPageElements(int pageNumber)
        {
            if (_info == null || _info.Slides == null)
                return new List<IElement>();

            var slide = _info.Slides.FirstOrDefault(s => s.SlideNumber == pageNumber);
            if (slide == null)
                return new List<IElement>();

            return slide.Elements.Cast<IElement>().ToList();
        }

        public List<IElement> GetAllElements()
        {
            if (_info == null || _info.Slides == null)
                return new List<IElement>();

            var result = new List<IElement>();
            foreach (var slide in _info.Slides)
            {
                result.AddRange(slide.Elements.Cast<IElement>());
            }
            return result;
        }

        public string ReadToJson(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return "{\"error\":\"文件路径不能为空\"}";

            if (!File.Exists(filePath))
                return $"{{\"error\":\"文件不存在: {filePath}\"}}";

            try
            {
                var info = _reader.ReadPresentation(filePath);
                return JsonConverter.ConvertToJson(info);
            }
            catch (Exception ex)
            {
                return $"{{\"error\":\"读取文件时出错: {ex.Message}\"}}";
            }
        }

        public PresentationInfo Read(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return new PresentationInfo { Error = "文件路径不能为空" };

            if (!File.Exists(filePath))
                return new PresentationInfo { Error = $"文件不存在: {filePath}" };

            try
            {
                return _reader.ReadPresentation(filePath);
            }
            catch (Exception ex)
            {
                return new PresentationInfo { Error = $"读取文件时出错: {ex.Message}" };
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _info = null;
                _disposed = true;
            }
        }
    }
}
