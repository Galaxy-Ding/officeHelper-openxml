using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class TextSoftEdgeJsonData
    {
        [JsonProperty("has_soft_edge")]
        public int HasSoftEdge { get; set; } = 0;
    }
}
