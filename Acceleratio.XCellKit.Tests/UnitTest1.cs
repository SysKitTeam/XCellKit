using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace Acceleratio.XCellKit.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void Test1()
        {
            var newExcel = new SpredsheetWorkbook();

            var worksheet = new SpredsheetWorksheet("Test22");
            worksheet.AddRow(new SpredsheetRow() { RowCells = new System.Collections.Generic.List<SpredsheetCell>() { new SpredsheetCell() { Value = "test" } } });

            newExcel.AddWorksheet(worksheet);
            newExcel.Save("c:\\temp\\test.xlsx");
        }
    }
}
