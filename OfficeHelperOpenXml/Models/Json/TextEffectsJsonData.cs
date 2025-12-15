using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TextEffectsJsonData
    {
        [JsonProperty("has_effects")]
        public int HasEffects { get; set; } = 0;

        [JsonProperty("shadow")]
        public TextShadowJsonData Shadow { get; set; }

        [JsonProperty("glow")]
        public TextGlowJsonData Glow { get; set; }

        [JsonProperty("reflection")]
        public TextReflectionJsonData Reflection { get; set; }

        [JsonProperty("soft_edge")]
        public TextSoftEdgeJsonData SoftEdge { get; set; }
    }
}
