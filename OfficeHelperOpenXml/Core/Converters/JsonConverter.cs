using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using OfficeHelperOpenXml.Core.Readers;

namespace OfficeHelperOpenXml.Core.Converters
{
    public static class JsonConverter
    {
        public static string ConvertToJson(PresentationInfo info)
        {
            if (info == null) return "{}";
            
            var sb = new StringBuilder();
            sb.Append("{\n");
            
            if (!string.IsNullOrEmpty(info.Error))
            {
                sb.Append($"\"error\":\"{EscapeJson(info.Error)}\",\n");
            }
            
            // 输出 master_slides
            sb.Append("\"master_slides\":[\n");
            for (int i = 0; i < info.MasterSlides.Count; i++)
            {
                if (i > 0) sb.Append(",\n");
                sb.Append(ConvertMasterSlideToJson(info.MasterSlides[i]));
            }
            sb.Append("\n],\n");
            
            // 输出 content_slides
            sb.Append("\"content_slides\":[\n");
            for (int i = 0; i < info.Slides.Count; i++)
            {
                if (i > 0) sb.Append(",\n");
                sb.Append(ConvertSlideToJson(info.Slides[i]));
            }
            sb.Append("\n]");
            sb.Append("\n}");
            
            return sb.ToString();
        }
        
        private static string ConvertSlideToJson(SlideInfo slide)
        {
            var sb = new StringBuilder();
            sb.Append("  {\n");
            sb.Append($"    \"page_number\":{slide.SlideNumber},\n");
            sb.Append($"    \"title\":\"\",\n");
            sb.Append($"    \"sub_title\":\"\",\n");
            sb.Append("    \"shapes\":[\n");
            
            for (int i = 0; i < slide.Elements.Count; i++)
            {
                if (i > 0) sb.Append(",\n");
                try
                {
                    sb.Append("      ");
                    sb.Append(slide.Elements[i].ToJson());
                }
                catch
                {
                    sb.Append("      {}");
                }
            }
            
            sb.Append("\n    ]");
            sb.Append("\n  }");
            
            return sb.ToString();
        }
        
        private static string ConvertMasterSlideToJson(MasterSlideInfo master)
        {
            var sb = new StringBuilder();
            sb.Append("  {\n");
            sb.Append($"    \"page_number\":{master.PageNumber},\n");
            sb.Append($"    \"title\":\"{EscapeJson(master.Title)}\",\n");
            sb.Append($"    \"sub_title\":\"{EscapeJson(master.SubTitle)}\",\n");
            sb.Append("    \"shapes\":[\n");
            
            for (int i = 0; i < master.Shapes.Count; i++)
            {
                if (i > 0) sb.Append(",\n");
                try
                {
                    sb.Append("      ");
                    sb.Append(master.Shapes[i].ToJson());
                }
                catch
                {
                    sb.Append("      {}");
                }
            }
            
            sb.Append("\n    ]");
            sb.Append("\n  }");
            
            return sb.ToString();
        }
        
        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
        }
    }
}
