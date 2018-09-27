using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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
        public void NonStreaming_LargeTable_FileCreated()
        {
            var newExcel = new SpreadsheetWorkbook();
            var columnsCount = 10;
            var worksheet = new SpreadsheetWorksheet("Test22");
            for (int i = 0; i < 100000; i++)
            {
                var cells = new List<SpreadsheetCell>();
                for (var j = 0; j < columnsCount; j++)
                {

                    cells.Add(new SpreadsheetCell()
                    {
                        BackgroundColor = Color.Red,
                        ForegroundColor = Color.Blue,
                        Font = _font,
                        Alignment = HorizontalAligment.Center,
                        Value = $"Ovo je test {i} - {j}"
                    });
                }

                var row = new SpreadsheetRow()
                {
                    RowCells = cells
                };
                worksheet.AddRow(row);
            }

            newExcel.AddWorksheet(worksheet);
            newExcel.Save(STR_TestOutputPath);

            Assert.IsTrue(File.Exists(STR_TestOutputPath));
        }

        [TestMethod]
        public void Streaming_LargeTable_MemoryConsumptionOk()
        {
            streaming_LargeTable_MemoryConsumptionOk(false);
        }

        [TestMethod]
        public void Streaming_LargeTableHyperLinks_MemoryConsumptionOk()
        {
            streaming_LargeTable_MemoryConsumptionOk(true);
        }

        private static void streaming_LargeTable_MemoryConsumptionOk(bool useHyperlinks)
        {
            var maxMemoryAllowed = useHyperlinks ? 150 : 40;

            var counter = 0;
            var maxMemDuringStreaming = 0.0;
            var newExcel = setupLargeWorkbook(row =>
            {
                counter++;
                if (counter % 1000 == 0)
                {
                    var mem = Utilities.GetMemoryConsumption();
                    if (mem > maxMemDuringStreaming)
                    {
                        maxMemDuringStreaming = mem;
                    }
                }
            }, useHyperLinks: useHyperlinks);

            var startingMemory = Utilities.GetMemoryConsumption();
            newExcel.Save(STR_TestOutputPath);
            var endingMemory = Utilities.GetMemoryConsumption();
            Console.WriteLine("Memory Consumption: ");
            Console.WriteLine("      Start: {0:N2}", startingMemory);
            Console.WriteLine("      End: {0:N2}", endingMemory);
            Console.WriteLine("      Max during streaming: {0:N2}", maxMemDuringStreaming);
            Assert.IsTrue(File.Exists(STR_TestOutputPath));

            
            Assert.IsTrue(endingMemory - startingMemory < maxMemoryAllowed, "Ending memory to high");
            Assert.IsTrue(maxMemDuringStreaming - startingMemory < maxMemoryAllowed, "Max memory to high");
        }

        [TestMethod]
        public void Streaming_LargeTable_TimeFactorOk()
        {
            var newExcel = setupLargeWorkbook(null);

            var sw = Stopwatch.StartNew();
            newExcel.Save(STR_TestOutputPath);
            sw.Stop();
            Assert.IsTrue(File.Exists(STR_TestOutputPath));
            Console.WriteLine("Export took: {0:N4} seconds", sw.Elapsed.TotalSeconds);
            Assert.IsTrue(sw.Elapsed.TotalSeconds < 60, "Export taking to long"); 
        }

        static Font _font = new Font(new FontFamily("Calibri"), 11);
        private static SpreadsheetWorkbook setupLargeWorkbook( Action<SpreadsheetRow> afterRowCreated, int rowsToStream = 800000, bool useHyperLinks = false)
        {             
            var columnsCount = 10;
            var newExcel = new SpreadsheetWorkbook();

            var worksheet = new SpreadsheetWorksheet("Test22");
            var table = new SpreadsheetTable("GridTable");
            for (var i = 0; i < 10; i++)
            {
                table.Columns.Add(new SpreadsheetTableColumn() {Name = $"Column{i}"});
            }

            table.ActivateStreamingMode();
            var rowCounter = 0;
            table.RequestTableRow += (s, args) =>
            {
                var cells = new List<SpreadsheetCell>();
                for (var i = 0; i < columnsCount; i++)
                {
                    if (useHyperLinks && i == columnsCount - 1)
                    {
                        cells.Add(new SpreadsheetHyperlinkCell(new SpreadsheetHyperLink($"http://www.google{rowCounter}.com", "google me!")));

                    }
                    else
                    {
                        cells.Add(new SpreadsheetCell()
                        {
                            BackgroundColor = Color.Red,
                            ForegroundColor = Color.Blue,
                            Font = _font,
                            Alignment = HorizontalAligment.Center,
                            Value = $"Ovo je test {rowCounter} - {i}"
                        });
                    }
                }

                args.Row = new SpreadsheetRow()
                {
                    RowCells = cells
                };
                rowCounter++;
                afterRowCreated?.Invoke(args.Row);
                args.Finished = rowCounter == rowsToStream;
            };

            worksheet.AddTable(table);

            newExcel.AddWorksheet(worksheet);
            return newExcel;
        }
    }
}
