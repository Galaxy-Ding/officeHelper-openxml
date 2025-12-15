using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TableCellJsonData
    {
        [JsonProperty("row")]
        public int Row { get; set; } = 0;

        [JsonProperty("col")]
        public int Col { get; set; } = 0;

        [JsonProperty("text")]
        public List<TextJsonData> Text { get; set; } = new List<TextJsonData>();

        [JsonProperty("hastext")]
        public int HasText { get; set; } = 0;

        [JsonProperty("fill")]
        public FillJsonData Fill { get; set; } = new FillJsonData();

        [JsonProperty("border")]
        public BorderJsonData Border { get; set; } = new BorderJsonData();

        [JsonProperty("box")]
        public string Box { get; set; } = "";

        [JsonProperty("merged")]
        public int Merged { get; set; } = 0;

        [JsonProperty("rowspan")]
        public int RowSpan { get; set; } = 1;

        [JsonProperty("colspan")]
        public int ColSpan { get; set; } = 1;

        [JsonProperty("text_align_horizontal")]
        public string TextAlignHorizontal { get; set; } = "left";

        [JsonProperty("text_align_vertical")]
        public string TextAlignVertical { get; set; } = "middle";

        [JsonProperty("text_orientation")]
        public string TextOrientation { get; set; } = "horizontal";

        [JsonProperty("text_rotation")]
        public float TextRotation { get; set; } = 0.0f;

        [JsonProperty("padding_top")]
        public float PaddingTop { get; set; } = 0.0f;

        [JsonProperty("padding_right")]
        public float PaddingRight { get; set; } = 0.0f;

        [JsonProperty("padding_bottom")]
        public float PaddingBottom { get; set; } = 0.0f;

        [JsonProperty("padding_left")]
        public float PaddingLeft { get; set; } = 0.0f;

        [JsonProperty("is_header")]
        public int IsHeader { get; set; } = 0;

        [JsonProperty("is_first_column")]
        public int IsFirstColumn { get; set; } = 0;
    }
}
