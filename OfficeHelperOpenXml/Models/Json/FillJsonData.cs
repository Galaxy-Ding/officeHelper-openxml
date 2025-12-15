using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class FillJsonData
    {
        [JsonProperty("color")]
        public string Color { get; set; } = "";

        [JsonProperty("opacity")]
        public float Opacity { get; set; } = 0;

        [JsonProperty("schemeColor")]
        public string SchemeColor { get; set; } = "";

        [JsonProperty("colorTransforms")]
        public ColorTransformJsonData ColorTransforms { get; set; }
    }
}
