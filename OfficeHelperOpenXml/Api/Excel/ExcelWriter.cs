using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeHelperOpenXml.Api.Excel
{
    public class ExcelWriter : IExcelWriter
    {
        private SpreadsheetDocument _document;
        private string _filePath;
        private bool _isOpen;
        private bool _disposed;

        public string FilePath => _filePath;
        public bool IsOpen => _isOpen;

        public bool OpenOrCreate(string filePath)
        {
            try
            {
                _document?.Dispose();
                if (File.Exists(filePath))
                {
                    _document = SpreadsheetDocument.Open(filePath, true);
                }
                else
                {
                    _document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
                    var workbookPart = _document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook(new Sheets());
                    workbookPart.Workbook.Save();
                    // Add default sheet directly here
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    worksheetPart.Worksheet.Save();
                    var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                    var sheet = new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    };
                    sheets.Append(sheet);
                    workbookPart.Workbook.Save();
                }
                _filePath = filePath;
                _isOpen = true;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OpenOrCreate error: {ex.Message}");
                _isOpen = false;
                return false;
            }
        }

        public bool CreateNew()
        {
            try
            {
                _document?.Dispose();
                var ms = new MemoryStream();
                _document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
                var workbookPart = _document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets());
                workbookPart.Workbook.Save();
                // Add default sheet
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                worksheetPart.Worksheet.Save();
                var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
                _filePath = null;
                _isOpen = true;
                return true;
            }
            catch { _isOpen = false; return false; }
        }

        public bool Save()
        {
            try
            {
                if (_document?.WorkbookPart?.Workbook != null)
                    _document.WorkbookPart.Workbook.Save();
                _document?.Save();
                return true;
            }
            catch { return false; }
        }

        public bool SaveAs(string filePath)
        {
            try
            {
                Save();
                _document?.Clone(filePath)?.Dispose();
                _filePath = filePath;
                return true;
            }
            catch { return false; }
        }

        public bool AddSheet(string sheetName)
        {
            if (_document == null) return false;
            try
            {
                var workbookPart = _document.WorkbookPart;
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                worksheetPart.Worksheet.Save();
                var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Any())
                    sheetId = sheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1;
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId,
                    Name = sheetName
                };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
                return true;
            }
            catch { return false; }
        }

        public bool DeleteSheet(string sheetName)
        {
            if (!_isOpen || _document == null) return false;
            try
            {
                var workbookPart = _document.WorkbookPart;
                var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                if (sheet == null) return false;
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                sheet.Remove();
                workbookPart.DeletePart(worksheetPart);
                workbookPart.Workbook.Save();
                return true;
            }
            catch { return false; }
        }

        public bool WriteCell(string sheetName, int row, int col, object value)
        {
            if (_document == null) return false;
            try
            {
                var worksheetPart = GetWorksheetPart(sheetName);
                if (worksheetPart == null) return false;
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var cellRef = GetCellReference(col, row);
                var rowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)row);
                if (rowElement == null)
                {
                    rowElement = new Row { RowIndex = (uint)row };
                    sheetData.Append(rowElement);
                }
                var cell = rowElement.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellRef);
                if (cell == null)
                {
                    cell = new Cell { CellReference = cellRef };
                    rowElement.Append(cell);
                }
                cell.CellValue = new CellValue(value?.ToString() ?? "");
                cell.DataType = CellValues.String;
                worksheetPart.Worksheet.Save();
                return true;
            }
            catch { return false; }
        }

        public bool WriteData(string sheetName, Dictionary<string, object> data)
        {
            return WriteData(sheetName, new List<Dictionary<string, object>> { data });
        }

        public bool WriteData(string sheetName, List<Dictionary<string, object>> dataList)
        {
            if (_document == null || dataList == null || dataList.Count == 0) return false;
            try
            {
                var worksheetPart = GetWorksheetPart(sheetName);
                if (worksheetPart == null)
                {
                    Console.WriteLine($"WriteData: Sheet '{sheetName}' not found");
                    return false;
                }
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                // Clear existing data
                sheetData.RemoveAllChildren<Row>();
                var headers = dataList[0].Keys.ToList();
                // Header row
                var headerRow = new Row { RowIndex = 1 };
                for (int i = 0; i < headers.Count; i++)
                {
                    var cell = new Cell { CellReference = GetCellReference(i, 1), CellValue = new CellValue(headers[i]), DataType = CellValues.String };
                    headerRow.Append(cell);
                }
                sheetData.Append(headerRow);
                // Data rows
                for (int r = 0; r < dataList.Count; r++)
                {
                    var dataRow = new Row { RowIndex = (uint)(r + 2) };
                    for (int c = 0; c < headers.Count; c++)
                    {
                        var val = dataList[r].ContainsKey(headers[c]) ? dataList[r][headers[c]]?.ToString() ?? "" : "";
                        var cell = new Cell { CellReference = GetCellReference(c, r + 2), CellValue = new CellValue(val), DataType = CellValues.String };
                        dataRow.Append(cell);
                    }
                    sheetData.Append(dataRow);
                }
                worksheetPart.Worksheet.Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WriteData error: {ex.Message}");
                return false;
            }
        }

        public List<string> GetSheetNames()
        {
            var names = new List<string>();
            if (_document?.WorkbookPart?.Workbook?.Sheets == null) return names;
            foreach (Sheet sheet in _document.WorkbookPart.Workbook.Sheets)
            {
                if (!string.IsNullOrEmpty(sheet.Name)) names.Add(sheet.Name);
            }
            return names;
        }

        private WorksheetPart GetWorksheetPart(string sheetName)
        {
            if (_document?.WorkbookPart?.Workbook?.Sheets == null) return null;
            var sheet = _document.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null) return null;
            return (WorksheetPart)_document.WorkbookPart.GetPartById(sheet.Id);
        }

        private string GetCellReference(int col, int row)
        {
            string colRef = "";
            col++;
            while (col > 0) { col--; colRef = (char)('A' + col % 26) + colRef; col /= 26; }
            return colRef + row;
        }

        public void Dispose()
        {
            if (_disposed) return;
            try { Save(); } catch { }
            _document?.Dispose();
            _disposed = true;
        }
    }
}
