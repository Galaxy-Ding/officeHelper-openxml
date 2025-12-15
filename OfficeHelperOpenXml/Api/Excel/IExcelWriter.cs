using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Api.Excel
{
    public interface IExcelWriter : IDisposable
    {
        string FilePath { get; }
        bool IsOpen { get; }
        bool OpenOrCreate(string filePath);
        bool CreateNew();
        bool Save();
        bool SaveAs(string filePath);
        bool AddSheet(string sheetName);
        bool DeleteSheet(string sheetName);
        bool WriteCell(string sheetName, int row, int col, object value);
        bool WriteData(string sheetName, Dictionary<string, object> data);
        bool WriteData(string sheetName, List<Dictionary<string, object>> dataList);
        List<string> GetSheetNames();
    }
}
