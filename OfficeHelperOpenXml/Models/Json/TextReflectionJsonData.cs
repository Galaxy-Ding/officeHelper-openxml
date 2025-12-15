using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TextReflectionJsonData
    {
        [JsonProperty("has_reflection")]
        public int HasReflection { get; set; } = 0;
    }
}
