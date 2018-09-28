using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit.SampleApp.Demos
{
    class LargeTableStreamingByEnumerator : LargeTableStreamingBase
    {
        public LargeTableStreamingByEnumerator() : base("Large Table Streaming by enumerator", "")
        {
        }

        public override void Execute()
        {
            var columnsCount = 10;
            var workBook = new SpreadsheetWorkbook();

            var worksheet = new SpreadsheetWorksheet("Test22");
            var table = new SpreadsheetTable("GridTable");
            for (var i = 0; i < 10; i++)
            {
                table.Columns.Add(new SpreadsheetTableColumn() { Name = $"Column{i}" });
            }

            var rowCounter = 0;
            var enumerator = Enumerable.Range(0, RowsToStream)
                // normally you would have some other IEnumerable that you would transform into spreadsheetrows
                // the key to use the least amount of memory as possible it to just pass it along without calling .ToList on the transformed data
                .Select(x =>
                {
                    return createTestRow(false, columnsCount, x);
                })
                .GetEnumerator();

            table.ActivateStreamingMode(enumerator);

            worksheet.AddTable(table);
            workBook.AddWorksheet(worksheet);
            workBook.Save(OutputFile);
        }
    }
}
