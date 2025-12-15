using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeHelperOpenXml.Api.Excel
{
    public class ExcelReader : IExcelReader
    {
        private SpreadsheetDocument _document;
        private string _filePath;
        private bool _isLoaded;
        private bool _disposed;

        public string FilePath => _filePath;
        public bool IsLoaded => _isLoaded;

        public bool Load(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                _document?.Dispose();
                _document = SpreadsheetDocument.Open(filePath, false);
                _filePath = filePath;
                _isLoaded = true;
                return true;
            }
            catch { _isLoaded = false; return false; }
        }

        public bool Reload()
        {
            if (string.IsNullOrEmpty(_filePath)) return false;
            return Load(_filePath);
        }

        public List<string> GetSheetNames()
        {
            var names = new List<string>();
            if (!_isLoaded || _document == null) return names;
            var workbookPart = _document.WorkbookPart;
            if (workbookPart?.Workbook?.Sheets == null) return names;
            foreach (Sheet sheet in workbookPart.Workbook.Sheets)
            {
                if (!string.IsNullOrEmpty(sheet.Name)) names.Add(sheet.Name);
            }
            return names;
        }

        public int GetSheetCount()
        {
            if (!_isLoaded || _document?.WorkbookPart?.Workbook?.Sheets == null) return 0;
            return _document.WorkbookPart.Workbook.Sheets.Count();
        }

        public List<Dictionary<string, object>> GetSheetData(string sheetName)
        {
            var result = new List<Dictionary<string, object>>();
            if (!_isLoaded || _document == null) return result;
            try
            {
                var workbookPart = _document.WorkbookPart;
                var sheet = workbookPart?.Workbook?.Sheets?.Elements<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);
                if (sheet == null) return result;
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var sheetData = worksheetPart?.Worksheet?.GetFirstChild<SheetData>();
                if (sheetData == null) return result;
                var rows = sheetData.Elements<Row>().ToList();
                if (rows.Count == 0) return result;
                var headers = new List<string>();
                var headerRow = rows[0];
                foreach (Cell cell in headerRow.Elements<Cell>())
                {
                    headers.Add(GetCellValue(cell, workbookPart));
                }
                for (int i = 1; i < rows.Count; i++)
                {
                    var rowData = new Dictionary<string, object>();
                    var cells = rows[i].Elements<Cell>().ToList();
                    for (int j = 0; j < headers.Count; j++)
                    {
                        var cellRef = GetCellReference(j, i + 1);
                        var cell = cells.FirstOrDefault(c => c.CellReference == cellRef);
                        rowData[headers[j]] = cell != null ? GetCellValue(cell, workbookPart) : "";
                    }
                    result.Add(rowData);
                }
            }
            catch { }
            return result;
        }

        public Dictionary<string, List<Dictionary<string, object>>> GetAllData()
        {
            var result = new Dictionary<string, List<Dictionary<string, object>>>();
            foreach (var name in GetSheetNames())
            {
                result[name] = GetSheetData(name);
            }
            return result;
        }

        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"sheetCount\":" + GetSheetCount() + ",");
            sb.Append("\"sheetNames\":[");
            var names = GetSheetNames();
            sb.Append(string.Join(",", names.Select(n => "\"" + EscapeJson(n) + "\"")));
            sb.Append("],\"sheets\":{");
            bool first = true;
            foreach (var name in names)
            {
                if (!first) sb.Append(",");
                first = false;
                sb.Append("\"" + EscapeJson(name) + "\":");
                sb.Append(SheetDataToJson(GetSheetData(name)));
            }
            sb.Append("}}");
            return sb.ToString();
        }

        public bool SaveToJson(string outputPath)
        {
            try { File.WriteAllText(outputPath, ToJson()); return true; }
            catch { return false; }
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell?.CellValue == null) return "";
            string value = cell.CellValue.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null && int.TryParse(value, out int index))
                {
                    var items = stringTable.SharedStringTable.Elements<SharedStringItem>().ToList();
                    if (index < items.Count) return items[index].InnerText;
                }
            }
            return value;
        }

        private string GetCellReference(int col, int row)
        {
            string colRef = "";
            col++;
            while (col > 0) { col--; colRef = (char)('A' + col % 26) + colRef; col /= 26; }
            return colRef + row;
        }

        private string SheetDataToJson(List<Dictionary<string, object>> data)
        {
            var sb = new StringBuilder();
            sb.Append("[");
            for (int i = 0; i < data.Count; i++)
            {
                if (i > 0) sb.Append(",");
                sb.Append("{");
                bool first = true;
                foreach (var kvp in data[i])
                {
                    if (!first) sb.Append(",");
                    first = false;
                    sb.Append("\"" + EscapeJson(kvp.Key) + "\":\"" + EscapeJson(kvp.Value?.ToString() ?? "") + "\"");
                }
                sb.Append("}");
            }
            sb.Append("]");
            return sb.ToString();
        }

        private string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
        }

        public void Dispose()
        {
            if (_disposed) return;
            _document?.Dispose();
            _disposed = true;
        }
    }
}
