using System;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeHelperOpenXml.Components;
using OfficeHelperOpenXml.Models;

namespace OfficeHelperOpenXml.Elements
{
    /// <summary>
    /// 表示连接线/箭头元素
    /// </summary>
    public class ConnectionElement : BaseElement
    {
        public override string ElementType => "connection";

        /// <summary>
        /// 位置和大小信息
        /// </summary>
        public ShapeBounds Bounds { get; set; }

        /// <summary>
        /// 起点连接的形状ID
        /// </summary>
        public string StartShapeId { get; set; } = "";

        /// <summary>
        /// 终点连接的形状ID
        /// </summary>
        public string EndShapeId { get; set; } = "";

        /// <summary>
        /// 起点连接索引
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// 终点连接索引
        /// </summary>
        public int EndIndex { get; set; }

        public ConnectionElement() : base()
        {
            Bounds = new ShapeBounds();
        }

        /// <summary>
        /// 实现基类的抽象方法（ConnectionElement 使用 InitializeFromConnectionShape 初始化）
        /// </summary>
        protected override void InitializeComponents(Shape shape, SlidePart slidePart)
        {
            // ConnectionElement 不使用 Shape 初始化，使用 InitializeFromConnectionShape 代替
        }

        /// <summary>
        /// 从 ConnectionShape 初始化
        /// </summary>
        public void InitializeFromConnectionShape(ConnectionShape connShape, SlidePart slidePart)
        {
            try
            {
                // 1. 提取名称和ID
                var nvPr = connShape.NonVisualConnectionShapeProperties;
                if (nvPr?.NonVisualDrawingProperties != null)
                {
                    Name = nvPr.NonVisualDrawingProperties.Name ?? "";
                }

                // 2. 提取位置和大小
                ExtractBounds(connShape);

                // 3. 提取连接信息
                ExtractConnectionInfo(nvPr);

                // 4. 提取线条样式（复用 LineComponent）
                var lineComponent = new LineComponent();
                var tempShape = CreateTempShape(connShape);
                lineComponent.ExtractFromShape(tempShape, slidePart);
                AddComponent(lineComponent);

                // 5. 提取填充（如果有）
                var fillComponent = new FillComponent();
                fillComponent.ExtractFromShape(tempShape, slidePart);
                AddComponent(fillComponent);
                
                // 6. 提取阴影（如果有）
                var shadowComponent = new ShadowComponent();
                shadowComponent.ExtractFromShape(tempShape, slidePart);
                AddComponent(shadowComponent);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"初始化 ConnectionElement 时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 提取位置和大小信息
        /// </summary>
        private void ExtractBounds(ConnectionShape connShape)
        {
            try
            {
                var spPr = connShape.ShapeProperties;
                if (spPr?.Transform2D != null)
                {
                    var xfrm = spPr.Transform2D;

                    long x = xfrm.Offset?.X ?? 0;
                    long y = xfrm.Offset?.Y ?? 0;
                    long cx = xfrm.Extents?.Cx ?? 0;
                    long cy = xfrm.Extents?.Cy ?? 0;
                    float rotation = xfrm.Rotation != null ? (float)(xfrm.Rotation.Value / 60000.0) : 0;

                    Bounds.SetFromEmu(x, y, cx, cy, rotation);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取 ConnectionShape 边界时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 提取连接信息
        /// </summary>
        private void ExtractConnectionInfo(NonVisualConnectionShapeProperties nvPr)
        {
            try
            {
                if (nvPr?.NonVisualConnectorShapeDrawingProperties != null)
                {
                    var cxnPr = nvPr.NonVisualConnectorShapeDrawingProperties;

                    // 提取起点连接
                    var startConn = cxnPr.GetFirstChild<A.StartConnection>();
                    if (startConn != null)
                    {
                        StartShapeId = startConn.Id?.Value.ToString() ?? "";
                        StartIndex = (int)(startConn.Index?.Value ?? 0);
                    }

                    // 提取终点连接
                    var endConn = cxnPr.GetFirstChild<A.EndConnection>();
                    if (endConn != null)
                    {
                        EndShapeId = endConn.Id?.Value.ToString() ?? "";
                        EndIndex = (int)(endConn.Index?.Value ?? 0);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取连接信息时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 创建临时 Shape 以复用现有组件
        /// </summary>
        private Shape CreateTempShape(ConnectionShape connShape)
        {
            var tempShape = new Shape();
            if (connShape.ShapeProperties != null)
            {
                tempShape.ShapeProperties = (ShapeProperties)connShape.ShapeProperties.CloneNode(true);
            }
            return tempShape;
        }

        /// <summary>
        /// 获取 Box 字符串（格式: "x,y,width,height"）
        /// </summary>
        private string Box => $"{Bounds.X:F2},{Bounds.Y:F2},{Bounds.Width:F2},{Bounds.Height:F2}";

        /// <summary>
        /// 转换为 JSON 字符串
        /// </summary>
        public override string ToJson()
        {
            var sb = new StringBuilder();
            sb.Append("{");

            // 基本信息
            sb.Append($"\"type\":\"{ElementType}\",");
            sb.Append($"\"name\":\"{EscapeJson(Name)}\",");
            sb.Append($"\"box\":\"{Box}\",");
            sb.Append($"\"rotation\":{Bounds.Rotation:F1},");

            // 连接信息（如果有）
            if (!string.IsNullOrEmpty(StartShapeId))
            {
                sb.Append($"\"start_shape_id\":\"{StartShapeId}\",");
                sb.Append($"\"start_index\":{StartIndex},");
            }
            if (!string.IsNullOrEmpty(EndShapeId))
            {
                sb.Append($"\"end_shape_id\":\"{EndShapeId}\",");
                sb.Append($"\"end_index\":{EndIndex},");
            }

            // 文本（连接线通常没有文本）
            sb.Append($"\"hastext\":0,");

            // 填充
            var fillComp = GetComponent<FillComponent>();
            if (fillComp != null)
            {
                sb.Append($"\"fill\":{fillComp.ToJson()},");
            }

            // 线条样式
            var lineComp = GetComponent<LineComponent>();
            if (lineComp != null)
            {
                sb.Append($"\"line\":{{{lineComp.ToJson()}}},");
            }
            
            // 阴影
            var shadowComp = GetComponent<ShadowComponent>();
            if (shadowComp != null)
            {
                var shadowJson = shadowComp.ToJson();
                if (shadowJson != "null")
                {
                    sb.Append($"\"shadow\":{{{shadowJson}}}");
                }
                else
                {
                    // 移除最后的逗号
                    if (sb[sb.Length - 1] == ',')
                    {
                        sb.Length--;
                    }
                }
            }
            else
            {
                // 移除最后的逗号
                if (sb[sb.Length - 1] == ',')
                {
                    sb.Length--;
                }
            }

            sb.Append("}");
            return sb.ToString();
        }

        /// <summary>
        /// 转义 JSON 字符串
        /// </summary>
        private string EscapeJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Replace("\\", "\\\\")
                      .Replace("\"", "\\\"")
                      .Replace("\n", "\\n")
                      .Replace("\r", "\\r")
                      .Replace("\t", "\\t");
        }
    }
}
