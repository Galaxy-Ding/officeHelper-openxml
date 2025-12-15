using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Api.Word
{
    public interface IWordReader : IDisposable
    {
        string FilePath { get; }
        bool IsLoaded { get; }
        bool Load(string filePath);
        bool Reload();
        string ToJson();
        bool SaveToJson(string outputPath);
        string GetFullText();
        List<string> GetParagraphs();
        List<Dictionary<string, object>> GetTables();
        int GetParagraphCount();
        int GetTableCount();
    }
}
