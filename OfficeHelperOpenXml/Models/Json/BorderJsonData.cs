using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class BorderJsonData
    {
        [JsonProperty("color")]
        public string Color { get; set; } = "";

        [JsonProperty("width")]
        public float Width { get; set; } = 0;
    }
}
