using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SysKit.XCellKit
{
    public class SpreadsheetTableColumn
    {
        public SpreadsheetTableColumn()
        {
            FilterValues = new HashSet<string>();
        }

        public string Name { get; set; }

        internal HashSet<string> FilterValues { get; private set; }

        public TableColumn ToOpenXmlColumn(int id)
        {
            return new TableColumn() {Id = new UInt32Value((UInt32)id), Name = Name};
        }

        public void AddFilterValue(string value)
        {
            if (!FilterValues.Contains(value))
            {
                FilterValues.Add(value);
            }
        }
    }
}
