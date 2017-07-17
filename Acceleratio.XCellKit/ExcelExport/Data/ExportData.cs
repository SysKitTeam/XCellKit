using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Acceleratio.XCellKit.ExcelExport
{
    public class ExportData
    {
        public static Stream OpenEntryStreamForWriting()
        {
            var filePath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "ExcelExport" , Guid.NewGuid().ToString() + ".xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            return File.Create(filePath);
        }

        /// <summary>
        /// Insert data to file using DocumentGenerator.
        /// </summary>
        public static void InsertData(Stream stream, DataTable data)
        {
            var spredsheetDocument = new SpredsheetWorkbook();
            var worksheet = new SpredsheetWorksheet("Results");

            var table = new SpredsheetTable("SearchResults");

            var columns = new List<SpredsheetTableColumn>();

            for (int i = 0; i < data.Columns.Count; i++)
            {
                columns.Add(new SpredsheetTableColumn() { Name = data.Columns[i].ColumnName });
            }

            table.Columns = columns;
            var tableRows = new List<SpredsheetRow>();
            for (int i = 0; i < data.Rows.Count; i++)
            {
                var row = new SpredsheetRow();// getTableRow(comapreResult[i], comapreResult.SourceTable.Rows[i], comapreResult.TargetTable.Rows[i], status, printObject);
                List<SpredsheetCell> cells = new List<SpredsheetCell>();
                for (int j = 0; j < data.Columns.Count; j++)
                {

                    var cell = new SpredsheetCell()
                    {
                        Value = data.Rows[i][j].ToString()
                    };
                    var dateType = SpredsheetDataTypeEnum.String;
                    //if (data.Columns[j].DataType == typeof(int) || data.Columns[j].DataType == typeof(float) || data.Columns[j].DataType == typeof(string))
                    //{
                    //    dateType = SpredsheetDataTypeEnum.Number;
                    //}
                    //else
                    if (data.Columns[j].DataType == typeof(DateTime))
                    {
                        dateType = SpredsheetDataTypeEnum.DateTime;
                    }

                    cell.SpredsheetDataType = dateType;
                    cells.Add(cell);

                }
                row.AddCellRange(cells);
                tableRows.Add(row);
            }

            table.Rows = tableRows;

            worksheet.AddTable(table, 1, 1);
            spredsheetDocument.AddWorksheet(worksheet);

            spredsheetDocument.Save(stream);
        }
    }
}
