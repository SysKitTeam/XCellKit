using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class SpredsheetHyperLink
    {
        public string Target { get; set; }
        public string DisplayValue { get; set; }

        public SpredsheetHyperLink(string target, string displayValue)
        {
            Target = target;
            DisplayValue = displayValue;
        }

        public SpredsheetHyperLink(SpredsheetWorksheet worksheet, SpredsheetLocation locationToFocus)
        {
            var worksheetName = new string(worksheet.Name.Take(30).ToArray());
            var locationRef = string.Format("{0}{1}", locationToFocus.ColumnName, locationToFocus.RowIndex);
            Target = string.Format("'{0}'!{1}", worksheetName, locationRef);
            DisplayValue = string.Format("Go to {0}.", worksheet.Name);
        }
        
        public void WriteHyperLink(OpenXmlWriter writer, int columnIndex, int rowIndex)
        {
            var refAtt = new OpenXmlAttribute("ref", null, string.Format("{0}{1}", SpredsheetHelper.ExcelColumnFromNumber(columnIndex), rowIndex));
            var locationAtt = new OpenXmlAttribute("location", null, Target);
            var displayAtt = new OpenXmlAttribute("display", null, DisplayValue);
            
            writer.WriteStartElement(new Hyperlink(), new List<OpenXmlAttribute>() {refAtt, locationAtt, displayAtt});
            writer.WriteEndElement();
        }
    }
}
