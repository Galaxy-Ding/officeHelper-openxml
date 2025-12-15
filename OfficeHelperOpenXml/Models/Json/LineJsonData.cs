using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class LineJsonData
    {
        [JsonProperty("has_outline")]
        public int HasOutline { get; set; } = 0;

        [JsonProperty("color")]
        public string Color { get; set; } = "";

        [JsonProperty("width")]
        public float Width { get; set; } = 0;

        [JsonProperty("schemeColor")]
        public string SchemeColor { get; set; } = "";
    }
}
