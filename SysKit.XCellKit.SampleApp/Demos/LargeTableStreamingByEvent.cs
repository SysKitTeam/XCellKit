using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit.SampleApp.Demos
{
    class LargeTableStreamingByEvent : LargeTableStreamingBase
    {
        public LargeTableStreamingByEvent() : base("Large Table Streaming by event", "")
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

            table.ActivateStreamingMode();
            var rowCounter = 0;
            table.TableRowRequested += (s, args) =>
            {
                var spreadsheetRow = createTestRow(true, columnsCount, rowCounter);
                args.Row = spreadsheetRow;
                rowCounter++;
                args.Finished = rowCounter == RowsToStream;
            };

            worksheet.AddTable(table);
            workBook.AddWorksheet(worksheet);
            workBook.Save(OutputFile);
        }
    }
}
