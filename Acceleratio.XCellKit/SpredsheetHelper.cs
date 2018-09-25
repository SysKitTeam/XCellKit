using System;
using System.Collections.Generic;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public static class SpreadsheetHelper
    {
        public static string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        public static SpreadsheetTable GetSpreadsheetTableFromDataTable(DataTable dataTable, string tableName)
        {
            var table = new SpreadsheetTable(tableName);
            table.Columns = new List<SpreadsheetTableColumn>();
            table.Rows = new List<SpreadsheetRow>();
            foreach (DataColumn dataColumn in dataTable.Columns)
            {
                table.Columns.Add(new SpreadsheetTableColumn() {Name = dataColumn.ColumnName});
            }

            foreach (DataRow dataRow in dataTable.Rows)
            {
                var row = new SpreadsheetRow();
                foreach (var o in dataRow.ItemArray)
                {
                    row.AddCell(new SpreadsheetCell() { Value = o });
                }

                table.Rows.Add(row);
            }

            return table;
        }

        public static HorizontalAlignmentValues GetHorizontalAlignmentValue(HorizontalAligment aligment)
        {
            switch (aligment)
            {
                case HorizontalAligment.General:
                    return HorizontalAlignmentValues.General;
                case HorizontalAligment.Left:
                    return HorizontalAlignmentValues.Left;
                case HorizontalAligment.Center:
                    return HorizontalAlignmentValues.Center;
                case HorizontalAligment.Right:
                    return HorizontalAlignmentValues.Right;
                case HorizontalAligment.Fill:
                    return HorizontalAlignmentValues.Fill;
                case HorizontalAligment.Justify:
                    return HorizontalAlignmentValues.Justify;
                case HorizontalAligment.CenterContinuous:
                    return HorizontalAlignmentValues.CenterContinuous;
                case HorizontalAligment.Distributed:
                    return HorizontalAlignmentValues.Distributed;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static SpreadsheetTable GetMonstrosity( string tableName)
        {
            var table = new SpreadsheetTable(tableName);
            table.Columns = new List<SpreadsheetTableColumn>();
            table.Rows = new List<SpreadsheetRow>();
            for(int i = 0; i< 10; i++)
            {
                table.Columns.Add(new SpreadsheetTableColumn() { Name = "Column" + i });
            }

            for(int i=0; i< 500000; i++)
            {
                var row = new SpreadsheetRow();
                for(int j = 0; j< 10; j++)
                {
                    row.AddCell(new SpreadsheetCell() { Value = string.Format("{0}-{1}", i, j) });
                }

                table.Rows.Add(row);
            }

            return table;
        }
    }
}
