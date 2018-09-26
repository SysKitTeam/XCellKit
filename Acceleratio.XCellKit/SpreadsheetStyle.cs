using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetStyle
    {
        public Color? BackgroundColor { get; set; }
        public Color? ForegroundColor { get; set; }
        public Font Font { get; set; }
        public HorizontalAlignmentValues? Alignment { get; set; }
        public VerticalAlignmentValues? VerticalAlignment { get; set; }
        public bool IsDate { get; set; }
        public int Indent { get; set; }
        public bool WrapText { get; set; }

        public string GetIdentifier()
        {
            var identifier = "";
            if (BackgroundColor.HasValue)
            {
                var colorRgb = BackgroundColor.Value.ToArgb();
                identifier += colorRgb;
            }
            if (ForegroundColor.HasValue)
            {
                var colorRgb = ForegroundColor.Value.ToArgb();
                identifier += colorRgb;
            }
            if (Font != null)
            {
                var fontid = Font.ToString();
                identifier += fontid;
            }
            if (Alignment.HasValue)
            {
                var aligment = Alignment.Value.ToString();
                identifier += aligment;
            }
            if (VerticalAlignment.HasValue)
            {
                var aligment = VerticalAlignment.Value.ToString();
                identifier += aligment;
            }
            if (IsDate)
            {
                identifier += IsDate;
            }

            identifier += WrapText;
            return identifier + Indent;
        }

        public override bool Equals(object obj)
        {
            var style = obj as SpreadsheetStyle;
            if (style == null)
            {
                return false;
            }

            return style.BackgroundColor == BackgroundColor && style.ForegroundColor == ForegroundColor && style.Font == Font && style.Alignment == Alignment && style.VerticalAlignment == VerticalAlignment && Indent == style.Indent && IsDate == style.IsDate;
        }

        public override int GetHashCode()
        {
            var hash = 0;
            if (BackgroundColor.HasValue)
            {
                var colorRgb = BackgroundColor.Value.ToArgb();
                hash ^= colorRgb.GetHashCode();
            }
            if (ForegroundColor.HasValue)
            {
                var colorRgb = ForegroundColor.Value.ToArgb();
                hash ^= colorRgb.GetHashCode();
            }
            if (Font != null)
            {
                var fontid = Font.ToString();
                hash ^= fontid.GetHashCode();
            }
            if (Alignment.HasValue)
            {
                var aligment = Alignment.Value.ToString();
                hash ^= aligment.GetHashCode();
            }

            if (VerticalAlignment.HasValue)
            {
                var aligment = VerticalAlignment.Value.ToString();
                hash ^= aligment.GetHashCode();
            }

            return hash ^ Indent ^ IsDate.GetHashCode() ^ WrapText.GetHashCode();
        }
    }
}
