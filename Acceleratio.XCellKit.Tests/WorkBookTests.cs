using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace Acceleratio.XCellKit.Tests
{
    [TestClass]
    public class WorkBookTests
    {
        const string STR_TestOutputPath = "C:\\temp\\test.xlsx";
        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(STR_TestOutputPath))
            {
                File.Delete(STR_TestOutputPath);
            }
        }

        [TestMethod]
        public void Save_SingleCell_FileCreated()
        {
            var newExcel = new SpreadsheetWorkbook();

            var worksheet = new SpreadsheetWorksheet("Test22");
            worksheet.AddRow(new SpreadsheetRow() { RowCells = new System.Collections.Generic.List<SpreadsheetCell>() { new SpreadsheetCell() { Value = "test" } } });

            newExcel.AddWorksheet(worksheet);
            newExcel.Save(STR_TestOutputPath);

            Assert.IsTrue(File.Exists(STR_TestOutputPath));
        }

        [TestMethod]
        public void Streaming_LargeTable_MemoryConsumptionOk()
        {
            var rowsToStream = 10;
            var columnsCount = 10;
            var newExcel = new SpreadsheetWorkbook();

            var worksheet = new SpreadsheetWorksheet("Test22");
            var table = new SpreadsheetTable("GridTable");
            for (var i = 0; i < 10; i++)
            {
                table.Columns.Add(new SpreadsheetTableColumn() { Name = $"Column{i}" });
            }

            table.ActivateStreamingMode(rowsToStream);
            var rowCounter = 0;
            table.RequestTableRow += (s, args) =>
            {
                var cells = new List<SpreadsheetCell>();
                for (var i = 0; i < columnsCount; i++)
                {
                    cells.Add(new SpreadsheetCell()
                    {
                        Value = $"Ovo je test {rowCounter} - {i}"
                    });
                }

                args.Row = new SpreadsheetRow()
                {
                    RowCells = cells
                };
                rowCounter++;
                args.Finished = rowCounter == rowsToStream;
            };

            worksheet.AddTable(table);

           // worksheet.AddRow(new SpreadsheetRow() { RowCells = new System.Collections.Generic.List<SpreadsheetCell>() { new SpreadsheetCell() { Value = "test" } } });

            newExcel.AddWorksheet(worksheet);
           
            newExcel.Save(STR_TestOutputPath);

            Assert.IsTrue(File.Exists(STR_TestOutputPath));
        }

        [TestMethod]
        public void Streaming_LargeTable_TimeFactorOk()
        {

        }


    }
}
