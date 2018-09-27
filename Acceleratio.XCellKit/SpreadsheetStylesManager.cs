using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public enum HorizontalAligment
    {
        General,
        Left,
        Center,
        Right,
        Fill,
        Justify,
        CenterContinuous,
        Distributed,
    }

    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom,
        Justify,
        Distributed,
    }

    public class SpreadsheetStylesManager
    {
        private Dictionary<string, int> _styles;
        private Dictionary<FontKey, int> _fonts;
        private Dictionary<System.Drawing.Color, int> _fills;
        private WorkbookPart _workbookPart;
        private Stylesheet _stylesheet;
        public SpreadsheetStylesManager(WorkbookPart workbookPart)
        {
            _workbookPart = workbookPart;
            _styles = new Dictionary<string, int>();
            _fonts = new Dictionary<FontKey, int>();
            _fills = new Dictionary<System.Drawing.Color, int>();

            var stylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
            _stylesheet = new Stylesheet();
            stylesPart.Stylesheet = _stylesheet;

            _stylesheet.Fills = new Fills();
            _stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            _stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            _stylesheet.Fills.Count = 2;

            stylesPart.Stylesheet.Fonts = new Fonts();
            stylesPart.Stylesheet.Fonts.Count = 2;
            stylesPart.Stylesheet.Fonts.AppendChild(new Font());

            Font hyperLinkFont = new Font();
            Underline underline1 = new Underline();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)10U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 238 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            hyperLinkFont.Append(underline1);
            hyperLinkFont.Append(fontSize2);
            hyperLinkFont.Append(color2);
            hyperLinkFont.Append(fontName2);
            hyperLinkFont.Append(fontFamilyNumbering2);
            hyperLinkFont.Append(fontCharSet2);
            hyperLinkFont.Append(fontScheme2);
            stylesPart.Stylesheet.Fonts.AppendChild(hyperLinkFont);

            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.Count = 1;
            stylesPart.Stylesheet.Borders.AppendChild(new Border());

            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellStyleFormats.Count = 2;
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            stylesPart.Stylesheet.CellFormats = new CellFormats();
            // empty one for index 0, seems to be required
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            CellFormat hyperLinkFormt = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            stylesPart.Stylesheet.CellFormats.AppendChild(hyperLinkFormt);
            stylesPart.Stylesheet.CellFormats.Count = 2;
            _hyperlinkStyles[new SpreadsheetStyle().GetIdentifier()] = 1;


            Borders borders = new Borders() { Count = (UInt32Value)1U };
            Border border = new Border();
            LeftBorder leftBorder = new LeftBorder();
            RightBorder rightBorder = new RightBorder();
            TopBorder topBorder = new TopBorder();
            BottomBorder bottomBorder = new BottomBorder();
            DiagonalBorder diagonalBorder = new DiagonalBorder();

            border.AppendChild(leftBorder);
            border.AppendChild(rightBorder);
            border.AppendChild(topBorder);
            border.AppendChild(bottomBorder);
            border.AppendChild(diagonalBorder);
            borders.AppendChild(border);

            _stylesheet.Borders = borders;

            CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles.AppendChild(cellStyle);
            _stylesheet.CellStyles = cellStyles;

            var dateTimeFormatInfo = DateTimeFormatInfo.CurrentInfo;
            var dateTimeFormat = "d.mm.yyyy hh:mm:ss";
            if (dateTimeFormatInfo != null)
            {
                dateTimeFormat = dateTimeFormatInfo.ShortDatePattern + " " + dateTimeFormatInfo.LongTimePattern;
                dateTimeFormat = dateTimeFormat.Replace("/", "\\/");
                dateTimeFormat = dateTimeFormat.Replace("tt", "AM/PM");
            }
            NumberingFormats numberingFormats = new NumberingFormats() { Count = (UInt32Value)1U };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)164U, FormatCode = dateTimeFormat };

            numberingFormats.Append(numberingFormat1);

            _stylesheet.NumberingFormats = numberingFormats;
        }

        private Dictionary<string, int> _hyperlinkStyles = new Dictionary<string, int>();
        public int GetHyperlinkStyleIndex(SpreadsheetStyle style)
        {
            var styleIdentifier = style.GetIdentifier();
            if (_hyperlinkStyles.ContainsKey(styleIdentifier))
            {
                return _hyperlinkStyles[styleIdentifier];
            }
            var cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            if (style.Alignment != null || style.Indent != 0 || style.VerticalAlignment != null)
            {
                var aligment = new Alignment();
                if (style.Alignment != null)
                {
                    aligment.Horizontal = style.Alignment.Value;
                }
                if (style.VerticalAlignment != null)
                {
                    aligment.Vertical = style.VerticalAlignment.Value;
                }
                if (style.Indent != 0)
                {
                    aligment.Indent = (UInt32)style.Indent;
                }
                cellFormat.AppendChild(aligment);
            }


            var styleIndex = _hyperlinkStyles[styleIdentifier] = (int)(UInt32)_stylesheet.CellFormats.Count;
            _stylesheet.CellFormats.AppendChild(cellFormat);
            _stylesheet.CellFormats.Count++;

            return styleIndex;

        }

        public int GetStyleIndex(SpreadsheetStyle style)
        {
            var styleIdentifier = style.GetIdentifier();
            if (_styles.ContainsKey(styleIdentifier))
            {
                return _styles[styleIdentifier];
            }

            var fontIndex = 0;
            if (style.Font != null)
            {
                var key = new FontKey(style.Font, style.ForegroundColor);
                if (_fonts.ContainsKey(key))
                {
                    fontIndex = _fonts[key];
                }
                else
                {
                    var newFont = new Font();
                    newFont.FontName = new FontName() { Val = style.Font.Name };
                    newFont.FontSize = new FontSize() { Val = style.Font.Size };
                    if (style.Font.Bold)
                    {
                        newFont.Bold = new Bold();
                    }
                    if (style.ForegroundColor != null)
                    {
                        newFont.Color = new Color() { Rgb = String.Format("{0:X2}{1:X2}{2:X2}{3:X2}", style.ForegroundColor.Value.A, style.ForegroundColor.Value.R, style.ForegroundColor.Value.G, style.ForegroundColor.Value.B) };
                    }
                    _stylesheet.Fonts.AppendChild(newFont);
                    fontIndex = _fonts[key] = (int)((UInt32)_stylesheet.Fonts.Count);
                    _stylesheet.Fonts.Count++;
                }
            }
            var fillIndex = 0;
            if (style.BackgroundColor != null)
            {
                if (_fills.ContainsKey(style.BackgroundColor.Value))
                {
                    fillIndex = _fills[style.BackgroundColor.Value];
                }
                else
                {
                    var newFill = new PatternFill() { PatternType = PatternValues.Solid };
                    newFill.ForegroundColor = new ForegroundColor() { Rgb = String.Format("{0:X2}{1:X2}{2:X2}{3:X2}", style.BackgroundColor.Value.A, style.BackgroundColor.Value.R, style.BackgroundColor.Value.G, style.BackgroundColor.Value.B) };
                    newFill.BackgroundColor = new BackgroundColor() { Indexed = 64U };
                    fillIndex = (int)(UInt32)_stylesheet.Fills.Count;
                    _fills[style.BackgroundColor.Value] = fillIndex;
                    _stylesheet.Fills.AppendChild(new Fill() { PatternFill = newFill });
                    _stylesheet.Fills.Count++;
                }
            }
            var numberingId = 0;
            if (style.IsDate)
            {
                numberingId = 164;
            }

            var cellFormat = new CellFormat() { FormatId = 0, FontId = (UInt32)fontIndex, FillId = (UInt32)fillIndex, BorderId = 0, NumberFormatId = (UInt32)numberingId, ApplyFill = true };
            if (style.Alignment != null || style.VerticalAlignment != null || style.Indent != 0 || style.WrapText)
            {
                var aligment = new Alignment();
                if (style.Alignment != null)
                {
                    aligment.Horizontal = style.Alignment.Value;
                }
                if (style.VerticalAlignment != null)
                {
                    aligment.Vertical = style.VerticalAlignment.Value;
                }
                if (style.Indent != 0)
                {
                    aligment.Indent = (UInt32)style.Indent;
                }
                if (style.WrapText)
                {
                    aligment.WrapText = true;
                }
                cellFormat.AppendChild(aligment);
            }

            var styleIndex = _styles[styleIdentifier] = (int)(UInt32)_stylesheet.CellFormats.Count;
            _stylesheet.CellFormats.AppendChild(cellFormat);
            _stylesheet.CellFormats.Count++;

            return styleIndex;
        }

        private class StyleKey
        {
            public StyleKey(int fontIndex, int fillIndex)
            {
                FontIndex = fontIndex;
                FillIndex = fillIndex;
            }
            public int FontIndex { get; private set; }
            public int FillIndex { get; private set; }

            public override bool Equals(object obj)
            {
                var styleKey = obj as StyleKey;
                if (styleKey == null)
                {
                    return false;
                }
                return styleKey.FontIndex == FontIndex && styleKey.FillIndex == FillIndex;
            }

            public override int GetHashCode()
            {
                return FontIndex.GetHashCode() ^ FillIndex.GetHashCode();
            }
        }

        public class FontKey : IEquatable<FontKey>
        {
            public FontKey(System.Drawing.Font font, System.Drawing.Color? color)
            {
                Font = font;
                Color = color;
            }

            public System.Drawing.Font Font { get; }
            public System.Drawing.Color? Color { get; }

            public bool Equals(FontKey other)
            {
                if (ReferenceEquals(null, other)) return false;
                if (ReferenceEquals(this, other)) return true;
                return Equals(Font, other.Font) && Color.Equals(other.Color);
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                if (ReferenceEquals(this, obj)) return true;
                if (obj.GetType() != this.GetType()) return false;
                return Equals((FontKey)obj);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    return ((Font != null ? Font.GetHashCode() : 0) * 397) ^ Color.GetHashCode();
                }
            }

            public static bool operator ==(FontKey left, FontKey right)
            {
                return Equals(left, right);
            }

            public static bool operator !=(FontKey left, FontKey right)
            {
                return !Equals(left, right);
            }
        }
    }


}
