using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeHelperOpenXml.Api.Word
{
    public class WordWriter : IWordWriter
    {
        private WordprocessingDocument _document;
        private string _filePath;
        private bool _isOpen;
        private bool _disposed;
        private const int EMU_PER_CM = 360000;

        public string FilePath => _filePath;
        public bool IsOpen => _isOpen;

        public bool OpenOrCreate(string filePath)
        {
            try
            {
                _document?.Dispose();
                if (File.Exists(filePath))
                {
                    _document = WordprocessingDocument.Open(filePath, true);
                }
                else
                {
                    _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
                    var mainPart = _document.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                }
                _filePath = filePath;
                _isOpen = true;
                return true;
            }
            catch { _isOpen = false; return false; }
        }

        public bool CreateNew()
        {
            try
            {
                _document?.Dispose();
                var ms = new MemoryStream();
                _document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
                var mainPart = _document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                _filePath = null;
                _isOpen = true;
                return true;
            }
            catch { _isOpen = false; return false; }
        }

        public bool Save()
        {
            try { _document?.Save(); return true; }
            catch { return false; }
        }

        public bool SaveAs(string filePath)
        {
            try
            {
                _document?.Clone(filePath)?.Dispose();
                _filePath = filePath;
                return true;
            }
            catch { return false; }
        }

        public bool AddParagraph(string text)
        {
            return AddParagraph(text, false, false, 12);
        }

        public bool AddParagraph(string text, bool isBold, bool isItalic, int fontSize)
        {
            if (!_isOpen || _document == null) return false;
            try
            {
                var body = _document.MainDocumentPart.Document.Body;
                var run = new Run(new Text(text));
                var runProps = new RunProperties();
                if (isBold) runProps.Append(new Bold());
                if (isItalic) runProps.Append(new Italic());
                runProps.Append(new FontSize { Val = (fontSize * 2).ToString() });
                run.PrependChild(runProps);
                var para = new Paragraph(run);
                body.Append(para);
                return true;
            }
            catch { return false; }
        }

        public bool AddHeading(string text, int level)
        {
            if (!_isOpen || _document == null || level < 1 || level > 9) return false;
            try
            {
                var body = _document.MainDocumentPart.Document.Body;
                var run = new Run(new Text(text));
                var runProps = new RunProperties(new Bold());
                int fontSize = level == 1 ? 32 : level == 2 ? 26 : level == 3 ? 24 : 22;
                runProps.Append(new FontSize { Val = (fontSize * 2).ToString() });
                run.PrependChild(runProps);
                var para = new Paragraph(run);
                var paraProps = new ParagraphProperties(new ParagraphStyleId { Val = "Heading" + level });
                para.PrependChild(paraProps);
                body.Append(para);
                return true;
            }
            catch { return false; }
        }

        public bool AddTable(List<List<string>> data)
        {
            if (!_isOpen || _document == null || data == null || data.Count == 0) return false;
            try
            {
                var body = _document.MainDocumentPart.Document.Body;
                var table = new Table();
                var tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                );
                table.Append(tblProps);
                foreach (var rowData in data)
                {
                    var row = new TableRow();
                    foreach (var cellText in rowData)
                    {
                        var cell = new TableCell(new Paragraph(new Run(new Text(cellText ?? ""))));
                        row.Append(cell);
                    }
                    table.Append(row);
                }
                body.Append(table);
                return true;
            }
            catch { return false; }
        }

        public bool InsertImage(string imagePath, double widthCm, double heightCm)
        {
            if (!_isOpen || _document == null || !File.Exists(imagePath)) return false;
            try
            {
                var mainPart = _document.MainDocumentPart;
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var stream = new FileStream(imagePath, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
                var relId = mainPart.GetIdOfPart(imagePart);
                long cx = (long)(widthCm * EMU_PER_CM);
                long cy = (long)(heightCm * EMU_PER_CM);
                var element = new Drawing(
                    new DW.Inline(
                        new DW.Extent { Cx = cx, Cy = cy },
                        new DW.DocProperties { Id = 1, Name = "Picture" },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties { Id = 0, Name = "Image" },
                                        new PIC.NonVisualPictureDrawingProperties()
                                    ),
                                    new PIC.BlipFill(
                                        new A.Blip { Embed = relId },
                                        new A.Stretch(new A.FillRectangle())
                                    ),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset { X = 0, Y = 0 },
                                            new A.Extents { Cx = cx, Cy = cy }
                                        ),
                                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                                    )
                                )
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    ) { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 }
                );
                var body = mainPart.Document.Body;
                body.Append(new Paragraph(new Run(element)));
                return true;
            }
            catch { return false; }
        }

        public bool InsertPageBreak()
        {
            if (!_isOpen || _document == null) return false;
            try
            {
                var body = _document.MainDocumentPart.Document.Body;
                body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                return true;
            }
            catch { return false; }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _document?.Dispose();
            _disposed = true;
        }
    }
}
