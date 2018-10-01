using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit.SampleApp.Demos
{
    class HelloWorldDemo: DemoBase
    {
        public HelloWorldDemo() : base("Hello World", "A small introduction")
        {
        }

        public override void Execute()
        {
            var workBook = new SpreadsheetWorkbook();
            var sheet = new SpreadsheetWorksheet("Sheet1");
            workBook.AddWorksheet(sheet);

            sheet.AddRow(new SpreadsheetRow()
            {
                RowCells = new List<SpreadsheetCell>()
                {
                    new SpreadsheetCell(){Value =  "Hello World!"}
                }
            });

            workBook.Save(OutputFile);
        }
    }
}
