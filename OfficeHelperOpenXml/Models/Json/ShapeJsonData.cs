using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeHelperOpenXml.Models.Json
{
    public class ShapeJsonData
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "";

        [JsonProperty("name")]
        public string Name { get; set; } = "";

        [JsonProperty("box")]
        public string Box { get; set; } = "";

        [JsonProperty("rotation")]
        public float Rotation { get; set; } = 0;

        [JsonProperty("fill")]
        public FillJsonData Fill { get; set; } = new FillJsonData();

        [JsonProperty("line")]
        public LineJsonData Line { get; set; } = new LineJsonData();

        [JsonProperty("shadow")]
        public ShadowJsonData Shadow { get; set; } = new ShadowJsonData();

        [JsonProperty("hastext")]
        public int HasText { get; set; } = 0;

        [JsonProperty("text")]
        public List<TextRunJsonData> Text { get; set; } = new List<TextRunJsonData>();

        /// <summary>
        /// Parses the box string (format: "left,top,width,height" in cm) into individual components
        /// </summary>
        /// <param name="left">Left position in cm</param>
        /// <param name="top">Top position in cm</param>
        /// <param name="width">Width in cm</param>
        /// <param name="height">Height in cm</param>
        /// <returns>True if parsing was successful, false otherwise</returns>
        public bool TryParseBox(out float left, out float top, out float width, out float height)
        {
            left = top = width = height = 0;

            if (string.IsNullOrEmpty(Box))
                return false;

            var parts = Box.Split(',');
            if (parts.Length != 4)
                return false;

            try
            {
                left = float.Parse(parts[0].Trim());
                top = float.Parse(parts[1].Trim());
                width = float.Parse(parts[2].Trim());
                height = float.Parse(parts[3].Trim());
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
