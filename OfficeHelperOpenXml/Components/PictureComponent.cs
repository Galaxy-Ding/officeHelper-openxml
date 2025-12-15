using System;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    public class PictureComponent : IElementComponent
    {
        public string ComponentType => "Picture";
        public bool IsEnabled { get; set; } = true;
        
        public bool HasPicture { get; set; }
        public string ImageBase64 { get; set; }
        public string ImageFormat { get; set; }
        public string ImagePath { get; set; }
        public int OriginalWidth { get; set; }
        public int OriginalHeight { get; set; }
        
        public PictureComponent()
        {
            HasPicture = false;
            ImageBase64 = "";
            ImageFormat = "";
            ImagePath = "";
        }
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
            HasPicture = false;
        }
        
        public void ExtractFromPicture(Picture picture, SlidePart slidePart)
        {
            if (slidePart == null) return;
            ExtractImageFromPart(picture, slidePart);
        }
        
        // 重载：从母版中提取图片
        public void ExtractFromPicture(Picture picture, SlideMasterPart masterPart)
        {
            if (masterPart == null) return;
            ExtractImageFromPart(picture, masterPart);
        }
        
        // 重载：从布局中提取图片
        public void ExtractFromPicture(Picture picture, SlideLayoutPart layoutPart)
        {
            if (layoutPart == null) return;
            ExtractImageFromPart(picture, layoutPart);
        }
        
        // 通用方法：从任意 Part 中提取图片
        private void ExtractImageFromPart(Picture picture, OpenXmlPart part)
        {
            try
            {
                HasPicture = false;
                var blipFill = picture.BlipFill;
                if (blipFill == null) return;
                
                var blip = blipFill.Blip;
                if (blip == null || blip.Embed == null) return;
                
                var relationshipId = blip.Embed.Value;
                var imagePart = part.GetPartById(relationshipId) as ImagePart;
                if (imagePart == null) return;
                
                ImageFormat = GetImageFormat(imagePart.ContentType);
                
                using (var stream = imagePart.GetStream())
                using (var ms = new MemoryStream())
                {
                    stream.CopyTo(ms);
                    var imageData = ms.ToArray();
                    ImageBase64 = Convert.ToBase64String(imageData);
                    
                    // Extract image dimensions
                    try
                    {
                        using (var imgStream = new MemoryStream(imageData))
                        {
                            var img = System.Drawing.Image.FromStream(imgStream);
                            OriginalWidth = img.Width;
                            OriginalHeight = img.Height;
                        }
                    }
                    catch
                    {
                        OriginalWidth = 0;
                        OriginalHeight = 0;
                    }
                }
                
                HasPicture = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取图片信息时出错: {ex.Message}");
                HasPicture = false;
            }
        }
        
        private string GetImageFormat(string contentType)
        {
            if (string.IsNullOrEmpty(contentType)) return "png";
            if (contentType.Contains("png")) return "png";
            if (contentType.Contains("jpeg") || contentType.Contains("jpg")) return "jpg";
            if (contentType.Contains("gif")) return "gif";
            if (contentType.Contains("bmp")) return "bmp";
            if (contentType.Contains("tiff")) return "tiff";
            if (contentType.Contains("wmf")) return "wmf";
            if (contentType.Contains("emf")) return "emf";
            return "png";
        }
        
        public void ApplyToShape(Shape shape, SlidePart slidePart) { }
        
        public Picture CreatePicture(SlidePart slidePart)
        {
            if (!HasPicture || string.IsNullOrEmpty(ImageBase64)) return null;
            
            try
            {
                var imageData = Convert.FromBase64String(ImageBase64);
                var contentType = GetContentType(ImageFormat);
                
                // 使用正确的方法添加图片部件
                ImagePart imagePart;
                if (ImageFormat == "png") imagePart = slidePart.AddImagePart("image/png");
                else if (ImageFormat == "jpg" || ImageFormat == "jpeg") imagePart = slidePart.AddImagePart("image/jpeg");
                else if (ImageFormat == "gif") imagePart = slidePart.AddImagePart("image/gif");
                else if (ImageFormat == "bmp") imagePart = slidePart.AddImagePart("image/bmp");
                else imagePart = slidePart.AddImagePart("image/png");
                
                using (var stream = new MemoryStream(imageData))
                {
                    imagePart.FeedData(stream);
                }
                
                var relationshipId = slidePart.GetIdOfPart(imagePart);
                
                var picture = new Picture();
                
                var nvPicPr = new NonVisualPictureProperties();
                nvPicPr.NonVisualDrawingProperties = new NonVisualDrawingProperties { Id = 1U, Name = "Picture" };
                nvPicPr.NonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
                nvPicPr.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
                picture.NonVisualPictureProperties = nvPicPr;
                
                var blipFill = new BlipFill();
                blipFill.Blip = new A.Blip { Embed = relationshipId };
                blipFill.Append(new A.Stretch(new A.FillRectangle()));
                picture.BlipFill = blipFill;
                
                picture.ShapeProperties = new ShapeProperties();
                
                return picture;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建图片时出错: {ex.Message}");
                return null;
            }
        }
        
        private string GetContentType(string format)
        {
            if (format == "png") return "image/png";
            if (format == "jpg" || format == "jpeg") return "image/jpeg";
            if (format == "gif") return "image/gif";
            if (format == "bmp") return "image/bmp";
            if (format == "tiff") return "image/tiff";
            return "image/png";
        }
        
        public string ToJson()
        {
            if (!IsEnabled) return "null";
            return $"\"imageFormat\":\"{ImageFormat.ToUpper()}\",\"imageWidth\":{OriginalWidth},\"imageHeight\":{OriginalHeight}";
        }

        public void FromJson(object jsonData)
        {
            try
            {
                if (jsonData == null) { HasPicture = false; return; }
                dynamic data = jsonData;
                try { HasPicture = Convert.ToInt32(data.hasPicture) != 0; } catch { HasPicture = false; }
                try { ImageFormat = data.imageFormat?.ToString() ?? "png"; } catch { ImageFormat = "png"; }
                try { ImageBase64 = data.imageBase64?.ToString() ?? ""; } catch { ImageBase64 = ""; }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从JSON加载图片信息时出错: {ex.Message}");
                HasPicture = false;
            }
        }
    }
}
