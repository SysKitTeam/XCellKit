using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit.SampleApp.Demos
{
    abstract class LargeTableStreamingBase : DemoBase
    {
        protected const int RowsToStream = 800000;
        static Font _font = new Font(new FontFamily("Calibri"), 11);
        public LargeTableStreamingBase(string title, string description) : base(title, description)
        {
        }

        protected SpreadsheetRow createTestRow(bool useHyperLinks, int columnsCount, int rowCounter)
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
                        Value = $"Cell value {rowCounter} - {i}"
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
