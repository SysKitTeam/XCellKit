using SkiaSharp;
using System.Collections.Generic;

namespace SysKit.XCellKit.SampleApp.Demos
{
    abstract class LargeTableStreamingBase : DemoBase
    {
        protected const int RowsToStream = 800000;
        private static SKFont _font = new SKFont(SKTypeface.FromFamilyName("Calibri"), 11);
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
                    cells.Add(new SpreadsheetHyperlinkCell(new SpreadsheetHyperLink($"http://www.google.com",
                        "google me!")));
                }
                else
                {
                    cells.Add(new SpreadsheetCell()
                    {
                        BackgroundColor = SKColor.Parse("#FF0000"),
                        ForegroundColor = SKColor.Parse("#0000FF"),
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
