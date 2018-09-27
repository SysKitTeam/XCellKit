using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetRow
    {
        public double? RowHeight { get; set; }
        public bool IsVisible { get; set; } = true;
        public SpreadsheetRow()
        {
            RowCells = new List<SpreadsheetCell>();
        }

        public void AddCell(SpreadsheetCell cell)
        {
            RowCells.Add(cell);
        }

        public void AddCellRange(List<SpreadsheetCell> cells)
        {
            RowCells.AddRange(cells);
        }


        public List<SpreadsheetCell> RowCells { get; set; }

        public void WriteRow(OpenXmlWriter writer, int columnIndex, int rowIndex, SpreadsheetStylesManager stylesManager, SpreadsheetHyperlinkManager hyperlinkManager, DrawingsManager drawingsManager)
        {
            var span = string.Format("{0}:{1}", columnIndex, RowCells.Count + columnIndex);
            var attributeList = new List<OpenXmlAttribute>();

            var rowIndexAtt = new OpenXmlAttribute("r", null, rowIndex.ToString());
            var spanAtt = new OpenXmlAttribute("spans", null, span);
            attributeList.Add(rowIndexAtt);
            attributeList.Add(spanAtt);
            if (this.RowHeight != null)
            {

                attributeList.Add(new OpenXmlAttribute("customHeight", null, "1"));
                attributeList.Add(new OpenXmlAttribute("ht", null, this.RowHeight.Value.ToString(CultureInfo.InvariantCulture)));
            }
            if (!IsVisible)
            {
                var hiddenAttribute = new OpenXmlAttribute("hidden", null, 1.ToString());
                attributeList.Add(hiddenAttribute);
            }

            writer.WriteStartElement(new Row(), attributeList);
            foreach (var cell in RowCells)
            {
                if (cell.ImageIndex != -1)
                {
                    drawingsManager.SetImageForCell(new ImageDetails()
                    {
                        Column = columnIndex,
                        ImageIndex = cell.ImageIndex,
                        ImageScaleFactor = cell.ImageScaleFactor,
                        Indent = cell.Indent,
                        Row = rowIndex
                    });
                }

                cell.WriteCell(writer, columnIndex, rowIndex, stylesManager, hyperlinkManager);
                columnIndex++;
            }
            writer.WriteEndElement();
        }
    }
}
