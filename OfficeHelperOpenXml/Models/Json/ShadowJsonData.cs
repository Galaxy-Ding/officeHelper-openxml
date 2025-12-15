using System;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class ShadowJsonData
    {
        [JsonProperty("has_shadow")]
        public int HasShadow { get; set; } = 0;

        [JsonProperty("color")]
        public string Color { get; set; } = "";

        [JsonProperty("opacity")]
        public float Opacity { get; set; } = 0;

        [JsonProperty("blur")]
        public float Blur { get; set; } = 0;

        [JsonProperty("offset_x")]
        public float OffsetX { get; set; } = 0;

        [JsonProperty("offset_y")]
        public float OffsetY { get; set; } = 0;

        [JsonProperty("size")]
        public float Size { get; set; } = 0;

        [JsonProperty("transparency")]
        public float Transparency { get; set; } = 0;

        [JsonProperty("type")]
        public object Type { get; set; } = 0;

        [JsonProperty("style")]
        public int Style { get; set; } = 0;

        [JsonProperty("shadow_type")]
        public string ShadowType { get; set; } = "";
    }
}
