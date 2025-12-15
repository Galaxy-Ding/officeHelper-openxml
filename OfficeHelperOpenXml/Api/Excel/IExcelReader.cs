using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Api.Excel
{
    public interface IExcelReader : IDisposable
    {
        string FilePath { get; }
        bool IsLoaded { get; }
        bool Load(string filePath);
        bool Reload();
        string ToJson();
        bool SaveToJson(string outputPath);
        List<Dictionary<string, object>> GetSheetData(string sheetName);
        Dictionary<string, List<Dictionary<string, object>>> GetAllData();
        List<string> GetSheetNames();
        int GetSheetCount();
    }
}
