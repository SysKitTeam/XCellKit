using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace Acceleratio.XCellKit.Tests
{
    [TestClass]
    public class WorkBookTests
    {
        [TestMethod]
        public void Save_SingleCell_FileCreated()
        {
            var newExcel = new SpreadsheetWorkbook();

            var worksheet = new SpreadsheetWorksheet("Test22");
            worksheet.AddRow(new SpreadsheetRow() { RowCells = new System.Collections.Generic.List<SpreadsheetCell>() { new SpreadsheetCell() { Value = "test" } } });

            newExcel.AddWorksheet(worksheet);
            var outputPath = "c:\\temp\\test.xlsx";
            newExcel.Save(outputPath);

            Assert.IsTrue(File.Exists(outputPath));
        }

        [TestMethod]
        public void Streaming_LargeTable_MemoryConsumptionOk()
        {

        }

        [TestMethod]
        public void Streaming_LargeTable_TimeFactorOk()
        {

        }


    }
}
