using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetTable
    {
        public SpreadsheetTable(string name)
        {
            Name = name;
            Columns = new List<SpreadsheetTableColumn>();
            Rows = new List<SpreadsheetRow>();
            FillLastCellInRow = true;
        }
        public string Name { get; private set; }
        public List<SpreadsheetTableColumn> Columns { get; set; }
        public List<SpreadsheetRow> Rows { get; set; }
        public bool FillLastCellInRow { get; set; }
        

        public Table GetTableDefinition(int id, int columnIndex, int rowIndex)
        {
            var startColumn = SpreadsheetHelper.ExcelColumnFromNumber(columnIndex);
            var endColumn = SpreadsheetHelper.ExcelColumnFromNumber(columnIndex + Columns.Count - 1);

            var reference = string.Format("{0}{1}:{2}{3}", startColumn, rowIndex, endColumn, rowIndex + Rows.Count);
            AutoFilter autoFilter = new AutoFilter() { Reference = reference };
            var table = new Table() {Id = (UInt32Value) (UInt32) id, Name = Name, DisplayName = Name.Replace(" ", "").Replace("-","") + id + "_", TotalsRowShown = false, Reference = reference};
            var tablesColumns = new TableColumns() {Count = (UInt32Value) (UInt32) Columns.Count};
            var i = 1;
            foreach (var column in Columns)
            {
                tablesColumns.Append(column.ToOpenXmlColumn(i));
                if (column.FilterValues.Any())
                {
                    var filterColumn = new FilterColumn() { ColumnId = (UInt32)i - 1};
                    var filters = new Filters();
                    foreach (var filterValue in column.FilterValues)
                    {
                        filters.AppendChild(new Filter() { Val = filterValue });
                    }
                    filterColumn.AppendChild(filters);
                    autoFilter.AppendChild(filterColumn);
                }
                i++;
            }
            TableStyleInfo tableStyle = new TableStyleInfo() { Name = "TableStyleLight9", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

            table.Append(autoFilter);
            table.Append(tablesColumns);
            table.Append(tableStyle);

            return table;

        }
    }
}
