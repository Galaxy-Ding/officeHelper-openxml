using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Api.Word
{
    public interface IWordWriter : IDisposable
    {
        string FilePath { get; }
        bool IsOpen { get; }
        bool OpenOrCreate(string filePath);
        bool CreateNew();
        bool Save();
        bool SaveAs(string filePath);
        bool AddParagraph(string text);
        bool AddParagraph(string text, bool isBold, bool isItalic, int fontSize);
        bool AddHeading(string text, int level);
        bool AddTable(List<List<string>> data);
        bool InsertImage(string imagePath, double widthCm, double heightCm);
        bool InsertPageBreak();
    }
}
