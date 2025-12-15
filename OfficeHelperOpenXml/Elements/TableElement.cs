using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;
using OfficeHelperOpenXml.Models.Json;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Elements
{
    public class TableElement : BaseElement
    {
        public override string ElementType => "table";
        
        private PositionComponent _positionComponent;
        private FillComponent _fillComponent;
        private LineComponent _lineComponent;
        private TableComponent _tableComponent;
        
        public TableElement()
        {
            _positionComponent = new PositionComponent();
            AddComponent(_positionComponent);
            
            _fillComponent = new FillComponent();
            AddComponent(_fillComponent);
            
            _lineComponent = new LineComponent();
            AddComponent(_lineComponent);
            
            _tableComponent = new TableComponent();
            AddComponent(_tableComponent);
        }
        
        protected override void InitializeComponents(Shape shape, SlidePart slidePart)
        {
        }
        
        public void ExtractFromGraphicFrame(GraphicFrame graphicFrame, SlidePart slidePart)
        {
            try
            {
                Name = graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";
                
                var transform = graphicFrame.Transform;
                if (transform != null)
                {
                    var offset = transform.Offset;
                    var extents = transform.Extents;
                    
                    if (offset != null && extents != null)
                    {
                        _positionComponent.Bounds.SetFromEmu(
                            offset.X?.Value ?? 0,
                            offset.Y?.Value ?? 0,
                            extents.Cx?.Value ?? 0,
                            extents.Cy?.Value ?? 0
                        );
                    }
                }
                
                _tableComponent.ExtractFromGraphicFrame(graphicFrame, slidePart);
            }
            catch (Exception)
            {
            }
        }
        
        public PositionComponent Position => _positionComponent;
        public FillComponent Fill => _fillComponent;
        public LineComponent Line => _lineComponent;
        public TableComponent Table => _tableComponent;
        public TableJsonData TableData => _tableComponent?.TableData;
        
        private int CheckTableHasText()
        {
            if (_tableComponent?.TableData == null)
                return 0;
            
            var tableData = _tableComponent.TableData;
            if (tableData.Cells == null || tableData.Cells.Count == 0)
                return 0;
            
            foreach (var row in tableData.Cells)
            {
                if (row == null) continue;
                
                foreach (var cell in row)
                {
                    if (cell == null) continue;
                    
                    if (cell.HasText == 1)
                    {
                        if (cell.Text != null && cell.Text.Count > 0)
                        {
                            foreach (var textItem in cell.Text)
                            {
                                if (textItem != null && !string.IsNullOrWhiteSpace(textItem.Content))
                                {
                                    return 1;
                                }
                            }
                        }
                    }
                }
            }
            
            return 0;
        }
        
        public override string ToJson()
        {
            var componentJsonParts = new List<string>();
            
            componentJsonParts.Add($"\"type\":\"{ElementType}\"");
            componentJsonParts.Add($"\"name\":\"{Name}\"");
            
            int hasText = CheckTableHasText();
            componentJsonParts.Add($"\"hastext\":{hasText}");
            
            if (_tableComponent != null)
            {
                var tableJson = _tableComponent.ToJson();
                if (tableJson != "null" && !string.IsNullOrWhiteSpace(tableJson))
                {
                    componentJsonParts.Add($"\"table\":{tableJson}");
                }
            }
            
            var nestedComponents = new Dictionary<string, string>
            {
                { "Fill", "fill" },
                { "Line", "line" }
            };
            
            foreach (var component in GetAllComponents().Where(c => c.IsEnabled && c.ComponentType != "Table"))
            {
                var componentJson = component.ToJson();
                if (componentJson != "null" && !string.IsNullOrWhiteSpace(componentJson))
                {
                    var content = componentJson.Trim();
                    
                    if (nestedComponents.TryGetValue(component.ComponentType, out var fieldName))
                    {
                        if (content.StartsWith("{") && content.EndsWith("}"))
                        {
                            content = content.Substring(1, content.Length - 2);
                        }
                        if (!string.IsNullOrWhiteSpace(content))
                        {
                            componentJsonParts.Add($"\"{fieldName}\":{{{content}}}");
                        }
                    }
                    else if (component.ComponentType == "Position")
                    {
                        if (content.StartsWith("{") && content.EndsWith("}"))
                        {
                            content = content.Substring(1, content.Length - 2);
                        }
                        if (!string.IsNullOrWhiteSpace(content))
                        {
                            componentJsonParts.Add(content);
                        }
                    }
                }
            }
            
            return "{" + string.Join(",", componentJsonParts) + "}";
        }
    }
}
