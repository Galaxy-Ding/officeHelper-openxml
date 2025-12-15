using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TextGlowJsonData
    {
        [JsonProperty("has_glow")]
        public int HasGlow { get; set; } = 0;
    }
}
