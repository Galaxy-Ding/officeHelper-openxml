using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeHelperOpenXml.Api.Word
{
    public class WordReader : IWordReader
    {
        private WordprocessingDocument _document;
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
                _document = WordprocessingDocument.Open(filePath, false);
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

        public string GetFullText()
        {
            if (!_isLoaded || _document == null) return "";
            try
            {
                var body = _document.MainDocumentPart?.Document?.Body;
                return body?.InnerText ?? "";
            }
            catch { return ""; }
        }

        public List<string> GetParagraphs()
        {
            var result = new List<string>();
            if (!_isLoaded || _document == null) return result;
            try
            {
                var body = _document.MainDocumentPart?.Document?.Body;
                if (body == null) return result;
                foreach (var para in body.Elements<Paragraph>())
                {
                    result.Add(para.InnerText);
                }
            }
            catch { }
            return result;
        }

        public List<Dictionary<string, object>> GetTables()
        {
            var result = new List<Dictionary<string, object>>();
            if (!_isLoaded || _document == null) return result;
            try
            {
                var body = _document.MainDocumentPart?.Document?.Body;
                if (body == null) return result;
                int tableIndex = 0;
                foreach (var table in body.Elements<Table>())
                {
                    var tableData = new Dictionary<string, object>();
                    tableData["index"] = tableIndex++;
                    var rows = new List<List<string>>();
                    foreach (var row in table.Elements<TableRow>())
                    {
                        var cells = new List<string>();
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            cells.Add(cell.InnerText);
                        }
                        rows.Add(cells);
                    }
                    tableData["rows"] = rows;
                    tableData["rowCount"] = rows.Count;
                    tableData["colCount"] = rows.Count > 0 ? rows[0].Count : 0;
                    result.Add(tableData);
                }
            }
            catch { }
            return result;
        }

        public int GetParagraphCount()
        {
            if (!_isLoaded || _document == null) return 0;
            try
            {
                return _document.MainDocumentPart?.Document?.Body?.Elements<Paragraph>().Count() ?? 0;
            }
            catch { return 0; }
        }

        public int GetTableCount()
        {
            if (!_isLoaded || _document == null) return 0;
            try
            {
                return _document.MainDocumentPart?.Document?.Body?.Elements<Table>().Count() ?? 0;
            }
            catch { return 0; }
        }

        public string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.Append("\"paragraphCount\":" + GetParagraphCount() + ",");
            sb.Append("\"tableCount\":" + GetTableCount() + ",");
            sb.Append("\"paragraphs\":[");
            var paras = GetParagraphs();
            sb.Append(string.Join(",", paras.Select(p => "\"" + EscapeJson(p) + "\"")));
            sb.Append("],\"tables\":[");
            var tables = GetTables();
            for (int i = 0; i < tables.Count; i++)
            {
                if (i > 0) sb.Append(",");
                sb.Append(TableToJson(tables[i]));
            }
            sb.Append("]}");
            return sb.ToString();
        }

        public bool SaveToJson(string outputPath)
        {
            try { File.WriteAllText(outputPath, ToJson()); return true; }
            catch { return false; }
        }

        private string TableToJson(Dictionary<string, object> table)
        {
            var sb = new StringBuilder();
            sb.Append("{\"index\":" + table["index"] + ",");
            sb.Append("\"rowCount\":" + table["rowCount"] + ",");
            sb.Append("\"colCount\":" + table["colCount"] + ",");
            sb.Append("\"rows\":[");
            var rows = (List<List<string>>)table["rows"];
            for (int r = 0; r < rows.Count; r++)
            {
                if (r > 0) sb.Append(",");
                sb.Append("[" + string.Join(",", rows[r].Select(c => "\"" + EscapeJson(c) + "\"")) + "]");
            }
            sb.Append("]}");
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
