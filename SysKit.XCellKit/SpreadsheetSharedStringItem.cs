using DocumentFormat.OpenXml.Spreadsheet;
using Color = SkiaSharp.SKColor;

namespace SysKit.XCellKit
{
    public class SpreadsheetSharedStringItem
    {
        public string Text { get; set; }
        public Color? FontColor { get; set; }

        public SpreadsheetSharedStringItem(string text, Color? fontColor = null)
        {
            Text = text;
            FontColor = fontColor;
        }

        internal SharedStringItem GetElement()
        {
            Run run = new Run();
            run.Append(new Text(Text));
            run.RunProperties = new RunProperties();

            if (FontColor.HasValue)
            {
                run.RunProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = FontColor.Value.GetRgbAsHex() });
            }

            return new SharedStringItem(run);
        }
    }
}
