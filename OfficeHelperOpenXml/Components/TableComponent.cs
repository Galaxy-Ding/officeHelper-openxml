using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeHelperOpenXml.Interfaces;
using OfficeHelperOpenXml.Models.Json;
using OfficeHelperOpenXml.Utils;
using Newtonsoft.Json;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeHelperOpenXml.Components
{
    public class TableComponent : IElementComponent
    {
        public string ComponentType => "Table";
        public bool IsEnabled { get; set; } = true;
        
        private TableJsonData _tableData;
        
        public TableComponent()
        {
            _tableData = new TableJsonData();
        }
        
        public TableJsonData TableData => _tableData;
        
        public void ExtractFromShape(Shape shape, SlidePart slidePart)
        {
        }
        
        public void ApplyToShape(Shape shape, SlidePart slidePart)
        {
        }
        
        public void FromJson(object jsonData)
        {
            if (jsonData == null) return;
            try
            {
                var json = jsonData.ToString();
                _tableData = JsonConvert.DeserializeObject<TableJsonData>(json);
            }
            catch { }
        }
        
        public void ExtractFromGraphicFrame(GraphicFrame graphicFrame, SlidePart slidePart)
        {
            try
            {
                var table = graphicFrame.Descendants<A.Table>().FirstOrDefault();
                if (table == null)
                {
                    _tableData = null;
                    return;
                }
                
                var rows = table.Descendants<A.TableRow>().ToList();
                int rowCount = rows.Count;
                int colCount = rows.FirstOrDefault()?.Descendants<A.TableCell>().Count() ?? 0;
                
                _tableData = new TableJsonData
                {
                    Rows = rowCount,
                    Columns = colCount,
                    Cells = new List<List<TableCellJsonData>>()
                };
                
                ExtractTableStyleInfo(table);
                
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var row = rows[rowIndex];
                    var cells = row.Descendants<A.TableCell>().ToList();
                    var rowCells = new List<TableCellJsonData>();
                    
                    for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                    {
                        var cell = cells[colIndex];
                        var cellData = ExtractCellData(cell, rowIndex, colIndex);
                        
                        if (rowIndex == 0 && _tableData.HasHeaderRow == 1)
                        {
                            cellData.IsHeader = 1;
                        }
                        if (colIndex == 0)
                        {
                            cellData.IsFirstColumn = 1;
                        }
                        
                        rowCells.Add(cellData);
                    }
                    
                    _tableData.Cells.Add(rowCells);
                }
            }
            catch (Exception)
            {
                _tableData = null;
            }
        }
        
        private void ExtractTableStyleInfo(A.Table table)
        {
            try
            {
                var tableProperties = table.TableProperties;
                if (tableProperties != null)
                {
                    _tableData.HasHeaderRow = tableProperties.FirstRow?.Value == true ? 1 : 0;
                    _tableData.HeaderRowIndex = _tableData.HasHeaderRow == 1 ? 0 : -1;
                    _tableData.HeaderRowHighlighted = _tableData.HasHeaderRow;
                    _tableData.FirstColumnHighlighted = tableProperties.FirstColumn?.Value == true ? 1 : 0;
                    _tableData.LastRowHighlighted = tableProperties.LastRow?.Value == true ? 1 : 0;
                    _tableData.LastColumnHighlighted = tableProperties.LastColumn?.Value == true ? 1 : 0;
                    _tableData.HasHorizontalBanding = tableProperties.BandRow?.Value == true ? 1 : 0;
                    _tableData.HasVerticalBanding = tableProperties.BandColumn?.Value == true ? 1 : 0;
                }
                _tableData.Level = 0;
            }
            catch { }
        }
        
        private TableCellJsonData ExtractCellData(A.TableCell cell, int rowIndex, int colIndex)
        {
            var cellData = new TableCellJsonData
            {
                Row = rowIndex,
                Col = colIndex,
                RowSpan = cell.RowSpan?.Value ?? 1,
                ColSpan = cell.GridSpan?.Value ?? 1,
                Merged = (cell.RowSpan?.Value > 1 || cell.GridSpan?.Value > 1) ? 1 : 0
            };
            
            try
            {
                var textBody = cell.TextBody;
                if (textBody != null)
                {
                    var paragraphs = textBody.Descendants<A.Paragraph>().ToList();
                    var textRuns = new List<TextJsonData>();
                    
                    foreach (var para in paragraphs)
                    {
                        foreach (var run in para.Descendants<A.Run>())
                        {
                            var text = run.Text?.Text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                var runJson = new TextJsonData { Content = text };
                                
                                var runProps = run.RunProperties;
                                if (runProps != null)
                                {
                                    runJson.FontSize = runProps.FontSize?.Value != null 
                                        ? runProps.FontSize.Value / 100f : 12f;
                                    runJson.FontBold = runProps.Bold?.Value == true ? 1 : 0;
                                    runJson.FontItalic = runProps.Italic?.Value == true ? 1 : 0;
                                    runJson.FontUnderline = runProps.Underline?.Value != null 
                                        && runProps.Underline.Value != A.TextUnderlineValues.None ? 1 : 0;
                                    
                                    var latin = runProps.GetFirstChild<A.LatinFont>();
                                    if (latin?.Typeface != null)
                                    {
                                        runJson.Font = latin.Typeface;
                                    }
                                    
                                    var solidFill = runProps.GetFirstChild<A.SolidFill>();
                                    if (solidFill != null)
                                    {
                                        runJson.FontColor = ColorHelper.ExtractColorString(solidFill);
                                    }
                                }
                                
                                textRuns.Add(runJson);
                            }
                        }
                    }
                    
                    if (textRuns.Count > 0)
                    {
                        cellData.HasText = 1;
                        cellData.Text = textRuns;
                    }
                    
                    var bodyProps = textBody.BodyProperties;
                    if (bodyProps != null)
                    {
                        var anchor = bodyProps.Anchor?.Value;
                        if (anchor == A.TextAnchoringTypeValues.Top)
                            cellData.TextAlignVertical = "top";
                        else if (anchor == A.TextAnchoringTypeValues.Center)
                            cellData.TextAlignVertical = "middle";
                        else if (anchor == A.TextAnchoringTypeValues.Bottom)
                            cellData.TextAlignVertical = "bottom";
                    }
                    
                    var firstPara = paragraphs.FirstOrDefault();
                    if (firstPara != null)
                    {
                        var paraProps = firstPara.ParagraphProperties;
                        if (paraProps?.Alignment != null)
                        {
                            var align = paraProps.Alignment.Value;
                            if (align == A.TextAlignmentTypeValues.Left)
                                cellData.TextAlignHorizontal = "left";
                            else if (align == A.TextAlignmentTypeValues.Center)
                                cellData.TextAlignHorizontal = "center";
                            else if (align == A.TextAlignmentTypeValues.Right)
                                cellData.TextAlignHorizontal = "right";
                            else if (align == A.TextAlignmentTypeValues.Justified)
                                cellData.TextAlignHorizontal = "justify";
                        }
                    }
                }
            }
            catch { }
            
            try
            {
                var cellProps = cell.TableCellProperties;
                if (cellProps != null)
                {
                    var solidFill = cellProps.GetFirstChild<A.SolidFill>();
                    if (solidFill != null)
                    {
                        cellData.Fill.Color = ColorHelper.ExtractColorString(solidFill);
                    }
                    
                    if (cellProps.LeftMargin?.Value != null)
                        cellData.PaddingLeft = (float)UnitConverter.EmuToCm(cellProps.LeftMargin.Value);
                    if (cellProps.RightMargin?.Value != null)
                        cellData.PaddingRight = (float)UnitConverter.EmuToCm(cellProps.RightMargin.Value);
                    if (cellProps.TopMargin?.Value != null)
                        cellData.PaddingTop = (float)UnitConverter.EmuToCm(cellProps.TopMargin.Value);
                    if (cellProps.BottomMargin?.Value != null)
                        cellData.PaddingBottom = (float)UnitConverter.EmuToCm(cellProps.BottomMargin.Value);
                }
            }
            catch { }
            
            return cellData;
        }
        
        public string ToJson()
        {
            if (!IsEnabled || _tableData == null)
                return "null";
            
            return JsonConvert.SerializeObject(_tableData, Formatting.None, new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                DefaultValueHandling = DefaultValueHandling.Include
            });
        }
    }
}
