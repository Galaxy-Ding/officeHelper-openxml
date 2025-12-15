using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class SlideJsonData
    {
        [JsonProperty("page_number")]
        public int PageNumber { get; set; } = 0;

        [JsonProperty("title")]
        public string Title { get; set; } = "";

        [JsonProperty("sub_title")]
        public string SubTitle { get; set; } = "";

        [JsonProperty("shapes")]
        public List<ShapeJsonData> Shapes { get; set; } = new List<ShapeJsonData>();
    }
}
