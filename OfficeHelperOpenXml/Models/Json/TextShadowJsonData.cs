using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TextShadowJsonData
    {
        [JsonProperty("has_shadow")]
        public int HasShadow { get; set; } = 0;

        [JsonProperty("type")]
        public string Type { get; set; } = "";

        [JsonProperty("color")]
        public string Color { get; set; } = "";

        [JsonProperty("blur")]
        public float Blur { get; set; } = 0;

        [JsonProperty("distance")]
        public float Distance { get; set; } = 0;

        [JsonProperty("angle")]
        public float Angle { get; set; } = 0;

        [JsonProperty("transparency")]
        public float Transparency { get; set; } = 0;
    }
}
