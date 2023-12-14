using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using Color = SkiaSharp.SKColor;

namespace SysKit.XCellKit
{
    public class SpreadsheetStyle
    {
        public Color? BackgroundColor { get; set; }
        public Color? ForegroundColor { get; set; }
        public SkiaSharp.SKFont Font { get; set; }
        public HorizontalAlignmentValues? Alignment { get; set; }
        public VerticalAlignmentValues? VerticalAlignment { get; set; }
        public bool IsDate { get; set; }
        public int Indent { get; set; }
        public bool WrapText { get; set; }

        public string GetIdentifier()
        {
            var sb = new StringBuilder();
            if (BackgroundColor.HasValue)
            {
                sb.Append(BackgroundColor.Value.GetRgbAsInt());
            }
            if (ForegroundColor.HasValue)
            {
                sb.Append(ForegroundColor.Value.GetRgbAsInt());
            }
            if (Font != null)
            {
                var fontid = Font.ToString();
                sb.Append(fontid);
                // sb.Append(Font.Style);
            }
            if (Alignment.HasValue)
            {
                // 30% speedup on 800 000 rows when not using Enum.ToString
                sb.Append("H");
                sb.Append((int)Alignment.Value);
            }
            if (VerticalAlignment.HasValue)
            {
                // 30% speedup on 800 000 rows when not using Enum.ToString
                sb.Append("V");
                sb.Append((int)VerticalAlignment.Value);
            }
            if (IsDate)
            {
                sb.Append(IsDate);
            }

            sb.Append(WrapText);
            sb.Append(Indent);

            return sb.ToString();
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
                hash ^= BackgroundColor.Value.GetRgbAsInt().GetHashCode();
            }
            if (ForegroundColor.HasValue)
            {
                hash ^= ForegroundColor.Value.GetRgbAsInt().GetHashCode();
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
