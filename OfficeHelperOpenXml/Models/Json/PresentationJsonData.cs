using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class PresentationJsonData
    {
        [JsonProperty("master_slides")]
        public List<SlideJsonData> MasterSlides { get; set; } = new List<SlideJsonData>();

        [JsonProperty("content_slides")]
        public List<SlideJsonData> ContentSlides { get; set; } = new List<SlideJsonData>();
    }
}
