using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Elements;

namespace OfficeHelperOpenXml.Core.Readers
{
    public class SlideReader
    {
        public List<BaseElement> ReadSlide(SlidePart slidePart)
        {
            var elements = new List<BaseElement>();

            if (slidePart?.Slide?.CommonSlideData?.ShapeTree == null)
                return elements;

            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            return ReadShapeTree(shapeTree, slidePart);
        }
        
        // 重载：读取母版
        public List<BaseElement> ReadSlide(SlideMasterPart masterPart)
        {
            var elements = new List<BaseElement>();

            if (masterPart?.SlideMaster?.CommonSlideData?.ShapeTree == null)
                return elements;

            var shapeTree = masterPart.SlideMaster.CommonSlideData.ShapeTree;
            return ReadShapeTreeFromMaster(shapeTree, masterPart);
        }
        
        // 重载：读取布局
        public List<BaseElement> ReadSlide(SlideLayoutPart layoutPart)
        {
            var elements = new List<BaseElement>();

            if (layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree == null)
                return elements;

            var shapeTree = layoutPart.SlideLayout.CommonSlideData.ShapeTree;
            return ReadShapeTreeFromLayout(shapeTree, layoutPart);
        }
        
        // 通用方法：读取 ShapeTree
        private List<BaseElement> ReadShapeTree(ShapeTree shapeTree, SlidePart slidePart)
        {
            var elements = new List<BaseElement>();
            
            foreach (var child in shapeTree.ChildElements)
            {
                try
                {
                    if (child is Shape shape)
                    {
                        var element = ElementFactory.CreateFromShape(shape, slidePart);
                        if (element != null) elements.Add(element);
                    }
                    else if (child is Picture picture)
                    {
                        var pictureElement = new PictureElement();
                        pictureElement.InitializeFromPicture(picture, slidePart);
                        elements.Add(pictureElement);
                    }
                    else if (child is GroupShape groupShape)
                    {
                        var groupElement = new GroupElement();
                        groupElement.InitializeFromGroupShape(groupShape, slidePart);
                        elements.Add(groupElement);
                    }
                    else if (child is GraphicFrame graphicFrame)
                    {
                        var element = ElementFactory.CreateFromGraphicFrame(graphicFrame, slidePart);
                        if (element != null) elements.Add(element);
                    }
                    else if (child is ConnectionShape connShape)
                    {
                        var connElement = new ConnectionElement();
                        connElement.InitializeFromConnectionShape(connShape, slidePart);
                        elements.Add(connElement);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"读取幻灯片元素时出错: {ex.Message}");
                }
            }

            return elements;
        }
        
        // 从母版读取 ShapeTree
        private List<BaseElement> ReadShapeTreeFromMaster(ShapeTree shapeTree, SlideMasterPart masterPart)
        {
            var elements = new List<BaseElement>();
            
            foreach (var child in shapeTree.ChildElements)
            {
                try
                {
                    if (child is Shape shape)
                    {
                        var element = ElementFactory.CreateFromShape(shape, null);
                        if (element != null) elements.Add(element);
                    }
                    else if (child is Picture picture)
                    {
                        var pictureElement = new PictureElement();
                        pictureElement.InitializeFromPicture(picture, masterPart);
                        elements.Add(pictureElement);
                    }
                    else if (child is GroupShape groupShape)
                    {
                        var groupElement = new GroupElement();
                        groupElement.InitializeFromGroupShape(groupShape, null);
                        elements.Add(groupElement);
                    }
                    else if (child is GraphicFrame graphicFrame)
                    {
                        var element = ElementFactory.CreateFromGraphicFrame(graphicFrame, null);
                        if (element != null) elements.Add(element);
                    }
                    else if (child is ConnectionShape connShape)
                    {
                        var connElement = new ConnectionElement();
                        connElement.InitializeFromConnectionShape(connShape, null);
                        elements.Add(connElement);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"读取母版元素时出错: {ex.Message}");
                }
            }

            return elements;
        }
        
        // 从布局读取 ShapeTree
        private List<BaseElement> ReadShapeTreeFromLayout(ShapeTree shapeTree, SlideLayoutPart layoutPart)
        {
            var elements = new List<BaseElement>();
            
            foreach (var child in shapeTree.ChildElements)
            {
                try
                {
                    if (child is Shape shape)
                    {
                        var element = ElementFactory.CreateFromShape(shape, null);
                        if (element != null) elements.Add(element);
                    }
                    else if (child is Picture picture)
                    {
                        var pictureElement = new PictureElement();
                        pictureElement.InitializeFromPicture(picture, layoutPart);
                        elements.Add(pictureElement);
                    }
                    else if (child is GroupShape groupShape)
                    {
                        var groupElement = new GroupElement();
                        groupElement.InitializeFromGroupShape(groupShape, null);
                        elements.Add(groupElement);
                    }
                    else if (child is GraphicFrame graphicFrame)
                    {
                        var element = ElementFactory.CreateFromGraphicFrame(graphicFrame, null);
                        if (element != null) elements.Add(element);
                    }
                    else if (child is ConnectionShape connShape)
                    {
                        var connElement = new ConnectionElement();
                        connElement.InitializeFromConnectionShape(connShape, null);
                        elements.Add(connElement);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"读取布局元素时出错: {ex.Message}");
                }
            }

            return elements;
        }
    }
}
