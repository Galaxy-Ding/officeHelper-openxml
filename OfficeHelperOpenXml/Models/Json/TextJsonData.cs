using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TextJsonData
    {
        [JsonProperty("content")]
        public string Content { get; set; } = "";

        [JsonProperty("font")]
        public string Font { get; set; } = "";

        [JsonProperty("font_size")]
        public float FontSize { get; set; } = 0;

        [JsonProperty("font_color")]
        public string FontColor { get; set; } = "";

        [JsonProperty("font_bold")]
        public int FontBold { get; set; } = 0;

        [JsonProperty("font_italic")]
        public int FontItalic { get; set; } = 0;

        [JsonProperty("font_underline")]
        public int FontUnderline { get; set; } = 0;

        [JsonProperty("font_strikethrough")]
        public int FontStrikethrough { get; set; } = 0;
    }
}
