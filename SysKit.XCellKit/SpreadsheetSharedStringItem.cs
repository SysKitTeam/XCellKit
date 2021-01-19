using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

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
                var hexColor = $"{FontColor.Value.R:X2}{FontColor.Value.G:X2}{FontColor.Value.B:X2}";
                run.RunProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = hexColor });
            }

            return new SharedStringItem(run);
        }
    }
}
