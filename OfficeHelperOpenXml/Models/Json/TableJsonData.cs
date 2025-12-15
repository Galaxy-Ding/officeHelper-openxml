using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TableJsonData
    {
        [JsonProperty("rows")]
        public int Rows { get; set; } = 0;

        [JsonProperty("columns")]
        public int Columns { get; set; } = 0;

        [JsonProperty("cells")]
        public List<List<TableCellJsonData>> Cells { get; set; } = new List<List<TableCellJsonData>>();

        [JsonProperty("style")]
        public string Style { get; set; } = "";

        [JsonProperty("level")]
        public int Level { get; set; } = 0;

        [JsonProperty("has_header_row")]
        public int HasHeaderRow { get; set; } = 0;

        [JsonProperty("header_row_index")]
        public int HeaderRowIndex { get; set; } = -1;

        [JsonProperty("first_column_highlighted")]
        public int FirstColumnHighlighted { get; set; } = 0;

        [JsonProperty("last_column_highlighted")]
        public int LastColumnHighlighted { get; set; } = 0;

        [JsonProperty("header_row_highlighted")]
        public int HeaderRowHighlighted { get; set; } = 0;

        [JsonProperty("last_row_highlighted")]
        public int LastRowHighlighted { get; set; } = 0;

        [JsonProperty("has_horizontal_banding")]
        public int HasHorizontalBanding { get; set; } = 0;

        [JsonProperty("has_vertical_banding")]
        public int HasVerticalBanding { get; set; } = 0;
    }
}
