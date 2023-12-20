namespace SysKit.XCellKit
{
    public class XCellFont
    {
        public string Name { get; set; }
        public int Size { get; set; }
        public XCellFontStyle Style { get; set; }

        public bool IsBold => Style == XCellFontStyle.Bold;
        public bool IsItalic => Style == XCellFontStyle.Italic;

        public XCellFont(string name, int size, XCellFontStyle style = XCellFontStyle.Regular)
        {
            Name = name;
            Size = size;
            Style = style;
        }

        public string GetId()
        {
            return $"[Font: Name={Name}, Size={Size}]";
        }
    }

    public enum XCellFontStyle
    {
        Regular = 0,
        Bold = 1,
        Italic = 2,
        Underline = 4,
        Strikeout = 8,
    }
}
