using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public static class SpreadsheetHelper
    {
        public static string[] columnNames;
        public static string ExcelColumnFromNumber(int column)
        {
            if (columnNames == null)
            {
                // bolje je unaprijed izgenerirati sva imena stupaca, nekih 8% je uzimalo na 500 000 ako se svaki put generiralo ispocetka
                pregenerateColumnNames();
            }
            if (column > columnNames.Length)
            {
                throw new InvalidOperationException($"No more than {columnNames.Length} columns supported.");
            }
            return columnNames[column];
        }

        private static void pregenerateColumnNames()
        {
            columnNames = new string[1024];
            for (int i = 1; i < 1024; i++)
            {
                columnNames[i] = excelColumnFromNumberCore(i);
            }
        }

        private static string excelColumnFromNumberCore(int column)
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
                table.Columns.Add(new SpreadsheetTableColumn() { Name = dataColumn.Caption });
            }

            foreach (DataRow dataRow in dataTable.Rows)
            {
                var row = new SpreadsheetRow();
                foreach (var o in dataRow.ItemArray)
                {
                    row.AddCell(new SpreadsheetCell() { Value = sanitizeString(o.ToString()) });
                }

                table.Rows.Add(row);
            }

            return table;
        }

        public static VerticalAlignmentValues GetVerticalAlignmentValues(VerticalAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalAlignment.Top:
                    return VerticalAlignmentValues.Top;
                case VerticalAlignment.Center:
                    return VerticalAlignmentValues.Center;
                case VerticalAlignment.Bottom:
                    return VerticalAlignmentValues.Bottom;
                case VerticalAlignment.Justify:
                    return VerticalAlignmentValues.Justify;
                case VerticalAlignment.Distributed:
                    return VerticalAlignmentValues.Distributed;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
            }
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
        public static ConditionalFormatValues GetConditionalFormatValues(ConditionalFormattingRuleTypeEnum type)
        {
            return (ConditionalFormatValues)(int)type;
        }

        public static SpreadsheetTable GetMonstrosity(string tableName)
        {
            var table = new SpreadsheetTable(tableName);
            table.Columns = new List<SpreadsheetTableColumn>();
            table.Rows = new List<SpreadsheetRow>();
            for (int i = 0; i < 10; i++)
            {
                table.Columns.Add(new SpreadsheetTableColumn() { Name = "Column" + i });
            }

            for (int i = 0; i < 500000; i++)
            {
                var row = new SpreadsheetRow();
                for (int j = 0; j < 10; j++)
                {
                    row.AddCell(new SpreadsheetCell() { Value = string.Format("{0}-{1}", i, j) });
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private static string sanitizeString(string s)
        {
            if (String.IsNullOrEmpty(s))
            {
                return s;
            }

            StringBuilder buffer = new StringBuilder(s.Length);

            for (int i = 0; i < s.Length; i++)
            {
                int code;
                try
                {
                    code = Char.ConvertToUtf32(s, i);
                }
                catch (ArgumentException)
                {
                    continue;
                }
                if (isLegalXmlChar(code))
                    buffer.Append(Char.ConvertFromUtf32(code));
                if (Char.IsSurrogatePair(s, i))
                    i++;
            }

            return buffer.ToString();
        }

        private static bool isLegalXmlChar(int codePoint)
        {
            return (codePoint == 0x9 ||
                codePoint == 0xA ||
                codePoint == 0xD ||
                (codePoint >= 0x20 && codePoint <= 0xD7FF) ||
                (codePoint >= 0xE000 && codePoint <= 0xFFFD) ||
                (codePoint >= 0x10000/* && character <= 0x10FFFF*/) //it's impossible to get a code point bigger than 0x10FFFF because Char.ConvertToUtf32 would have thrown an exception
            );
        }
    }
}
