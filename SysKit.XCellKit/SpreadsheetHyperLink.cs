using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SysKit.XCellKit
{
    public class SpreadsheetHyperLink
    {
        public string Target { get; set; }
        public string DisplayValue { get; set; }

        public SpreadsheetHyperLink(string target, string displayValue)
        {
            Target = target;
            DisplayValue = displayValue;
        }

        public SpreadsheetHyperLink(SpreadsheetWorksheet worksheet, SpreadsheetLocation locationToFocus)
        {
            var worksheetName = new string(worksheet.Name.Take(30).ToArray());
            var locationRef = string.Format("{0}{1}", locationToFocus.ColumnName, locationToFocus.RowIndex);
            Target = string.Format("'{0}'!{1}", worksheetName, locationRef);
            DisplayValue = string.Format("Go to {0}.", worksheet.Name);
        }
    }
}
