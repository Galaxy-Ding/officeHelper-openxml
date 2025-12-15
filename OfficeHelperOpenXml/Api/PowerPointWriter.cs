using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Elements;
using OfficeHelperOpenXml.Core.Writers;
using OfficeHelperOpenXml.Core.Converters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeHelperOpenXml.Api
{
    /// <summary>
    /// PowerPointд���� - ����OpenXMLʵ��
    /// </summary>
    public class PowerPointWriter : IPowerPointWriter
    {
        private PresentationDocument _document;
        private string _filePath;
        private bool _disposed;
        private bool _isNewFile;
        private SlideWriter _slideWriter;
        private RelationshipIdManager _relationshipIdManager;

        public string FilePath => _filePath;
        public bool IsOpen => _document != null;

        public bool OpenOrCreate(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return false;

            try
            {
                _filePath = filePath;
                
                // Initialize RelationshipIdManager and SlideWriter
                _relationshipIdManager = new RelationshipIdManager();
                _slideWriter = new SlideWriter(_relationshipIdManager);
                
                if (File.Exists(filePath))
                {
                    _document = PresentationDocument.Open(filePath, true);
                    _isNewFile = false;
                }
                else
                {
                    var dir = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    
                    _document = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
                    _isNewFile = true;
                    InitializePresentation();
                }
                return true;
            }
            catch { return false; }
        }

        public bool CreateNew()
        {
            try
            {
                var tempPath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.pptx");
                _document = PresentationDocument.Create(tempPath, PresentationDocumentType.Presentation);
                _filePath = tempPath;
                _isNewFile = true;
                InitializePresentation();
                return true;
            }
            catch { return false; }
        }

        private void InitializePresentation()
        {
            var presentationPart = _document.AddPresentationPart();
            presentationPart.Presentation = new Presentation(
                new SlideIdList(),
                new SlideSize { Cx = 12192000, Cy = 6858000 },
                new NotesSize { Cx = 6858000, Cy = 9144000 }
            );
            
            // Add slide master
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = CreateSlideMaster();
            
            // Add slide layout
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = CreateSlideLayout();
            
            // Link layout to master
            slideMasterPart.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                new SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }
            );
            
            // Add master to presentation
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                new SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }
            );
            
            AddSlide();
        }

        private SlideMaster CreateSlideMaster()
        {
            return new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup())
                ))
            );
        }

        private SlideLayout CreateSlideLayout()
        {
            return new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup())
                ))
            ) { Type = SlideLayoutValues.Blank };
        }

        public bool Save()
        {
            if (!IsOpen) return false;
            try { _document.Save(); return true; }
            catch { return false; }
        }

        public bool SaveAs(string filePath)
        {
            if (!IsOpen || string.IsNullOrEmpty(filePath)) return false;
            try
            {
                var dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                
                _document.Clone(filePath).Dispose();
                _document = PresentationDocument.Open(filePath, true);
                _filePath = filePath;
                return true;
            }
            catch { return false; }
        }

        public int AddSlide()
        {
            if (!IsOpen) return -1;
            try
            {
                var presentationPart = _document.PresentationPart;
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                
                slidePart.Slide = new Slide(
                    new CommonSlideData(new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup())
                    ))
                );

                // Link to layout
                var masterPart = presentationPart.SlideMasterParts.FirstOrDefault();
                if (masterPart != null)
                {
                    var layoutPart = masterPart.SlideLayoutParts.FirstOrDefault();
                    if (layoutPart != null)
                        slidePart.AddPart(layoutPart);
                }

                var slideIdList = presentationPart.Presentation.SlideIdList;
                uint maxSlideId = slideIdList.ChildElements.Count > 0
                    ? slideIdList.ChildElements.Cast<SlideId>().Max(s => s.Id.Value)
                    : 255;

                var slideId = new SlideId
                {
                    Id = maxSlideId + 1,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };
                slideIdList.Append(slideId);

                return slideIdList.ChildElements.Count;
            }
            catch { return -1; }
        }

        public bool DeleteSlide(int slideIndex)
        {
            if (!IsOpen || slideIndex < 1) return false;
            try
            {
                var presentationPart = _document.PresentationPart;
                var slideIdList = presentationPart.Presentation.SlideIdList;
                
                if (slideIndex > slideIdList.ChildElements.Count) return false;
                
                var slideId = slideIdList.ChildElements.ElementAt(slideIndex - 1) as SlideId;
                if (slideId == null) return false;
                
                var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                slideIdList.RemoveChild(slideId);
                
                if (slidePart != null)
                    presentationPart.DeletePart(slidePart);
                
                return true;
            }
            catch { return false; }
        }

        public bool AddElement(int slideIndex, IElement element)
        {
            if (!IsOpen || element == null || slideIndex < 1) return false;
            try
            {
                var slidePart = GetSlidePart(slideIndex);
                if (slidePart == null) return false;
                
                return _slideWriter.AddElement(slidePart, element);
            }
            catch { return false; }
        }

        public bool AddElements(int slideIndex, IEnumerable<IElement> elements)
        {
            if (!IsOpen || elements == null) return false;
            bool success = true;
            foreach (var element in elements)
            {
                if (!AddElement(slideIndex, element))
                    success = false;
            }
            return success;
        }

        public bool UpdateElement(int slideIndex, int elementIndex, IElement element)
        {
            if (!IsOpen || element == null) return false;
            try
            {
                var slidePart = GetSlidePart(slideIndex);
                if (slidePart == null) return false;
                
                return _slideWriter.UpdateElement(slidePart, elementIndex, element);
            }
            catch { return false; }
        }

        public bool DeleteElement(int slideIndex, int elementIndex)
        {
            if (!IsOpen) return false;
            try
            {
                var slidePart = GetSlidePart(slideIndex);
                if (slidePart == null) return false;
                
                return _slideWriter.DeleteElement(slidePart, elementIndex);
            }
            catch { return false; }
        }

        public int GetSlideCount()
        {
            if (!IsOpen) return 0;
            try { return _document.PresentationPart.Presentation.SlideIdList.ChildElements.Count; }
            catch { return 0; }
        }

        private SlidePart GetSlidePart(int slideIndex)
        {
            if (slideIndex < 1) return null;
            var slideIdList = _document.PresentationPart.Presentation.SlideIdList;
            if (slideIndex > slideIdList.ChildElements.Count) return null;
            
            var slideId = slideIdList.ChildElements.ElementAt(slideIndex - 1) as SlideId;
            return _document.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _document?.Dispose();
                _document = null;
                _disposed = true;
            }
        }
    }
}
