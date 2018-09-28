using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit.SampleApp.Demos
{
    class SimpleTableDemo : DemoBase
    {
        public SimpleTableDemo() : base("Simple table", "")
        {
        }

        public override void Execute()
        {
            var workBook = new SpreadsheetWorkbook();
            var sheet = new SpreadsheetWorksheet("Sheet1");
            workBook.AddWorksheet(sheet);

            var table = new SpreadsheetTable("Table1");
           
            table.Columns = new List<SpreadsheetTableColumn>()
            {
                new SpreadsheetTableColumn()
                {
                    Name = "Column A"
                },
                new SpreadsheetTableColumn()
                {
                    Name = "Column B"
                }
            };

            table.Rows = new List<SpreadsheetRow>()
            {
                new SpreadsheetRow()
                {
                    RowCells = new List<SpreadsheetCell>()
                    {
                        new SpreadsheetCell()
                        {
                            Value = "Test"
                        },
                        new SpreadsheetHyperlinkCell(new SpreadsheetHyperLink("https://www.google.com", "Google me!"))
                    }
                },
                new SpreadsheetRow()
                {
                    RowCells = new List<SpreadsheetCell>()
                    {
                        new SpreadsheetCell()
                        {
                            Value = "Test2 "
                        },
                        new SpreadsheetHyperlinkCell(new SpreadsheetHyperLink("https://www.google.com", "Google me!"))
                    }
                }
            };

            sheet.AddTable(table);


            workBook.Save(OutputFile);
        }
    }
}
