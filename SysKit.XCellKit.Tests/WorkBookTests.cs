using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace SysKit.XCellKit.Tests
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

        [TestMethod]
        public void StreamingEnumerator_LargeTable_MemoryConsumptionOk()
        {
            streaming_LargeTable_MemoryConsumptionOk(false, true);
        }

        private static void streaming_LargeTable_MemoryConsumptionOk(bool useHyperlinks, bool useEnumerator = false)
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
            }, useHyperLinks: useHyperlinks, useEnumerator: useEnumerator);

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
        private static SpreadsheetWorkbook setupLargeWorkbook( Action<SpreadsheetRow> afterRowCreated, int rowsToStream = 800000, bool useHyperLinks = false, bool useEnumerator = false)
        {             
            var columnsCount = 10;
            var newExcel = new SpreadsheetWorkbook();

            var worksheet = new SpreadsheetWorksheet("Test22");
            var table = new SpreadsheetTable("GridTable");
            for (var i = 0; i < 10; i++)
            {
                table.Columns.Add(new SpreadsheetTableColumn() {Name = $"Column{i}"});
            }

            if (!useEnumerator)
            {
                table.ActivateStreamingMode();
                var rowCounter = 0;
                table.TableRowRequested += (s, args) =>
                {
                    var spreadsheetRow = createTestRow(useHyperLinks, columnsCount, rowCounter);
                    args.Row = spreadsheetRow;
                    rowCounter++;
                    afterRowCreated?.Invoke(args.Row);
                    args.Finished = rowCounter == rowsToStream;
                };
            }
            else
            {
                var enumerator = Enumerable.Range(0, rowsToStream).Select(x =>
                    {
                        var row = createTestRow(useHyperLinks, columnsCount, x);
                        afterRowCreated?.Invoke(row);                        
                        return row;
                    })
                    .GetEnumerator();

                table.ActivateStreamingMode(enumerator);
            }

            worksheet.AddTable(table);

            newExcel.AddWorksheet(worksheet);
            return newExcel;
        }

        private static SpreadsheetRow createTestRow(bool useHyperLinks, int columnsCount, int rowCounter)
        {
            var cells = new List<SpreadsheetCell>();
            for (var i = 0; i < columnsCount; i++)
            {
                if (useHyperLinks && i == columnsCount - 1)
                {
                    cells.Add(new SpreadsheetHyperlinkCell(new SpreadsheetHyperLink($"http://www.google{rowCounter}.com",
                        "google me!")));
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

            var spreadsheetRow = new SpreadsheetRow()
            {
                RowCells = cells
            };
            return spreadsheetRow;
        }
    }
}
