using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace SysKit.XCellKit
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
        private Dictionary<SixLabors.ImageSharp.Color, int> _fills;

        private Stylesheet _stylesheet;
        public SpreadsheetStylesManager()
        {
            _styles = new Dictionary<string, int>();
            _fonts = new Dictionary<FontKey, int>();
            _fills = new Dictionary<SixLabors.ImageSharp.Color, int>();


            _stylesheet = new Stylesheet();

            _stylesheet.Fills = new Fills();
            _stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            _stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            _stylesheet.Fills.Count = 2;

            _stylesheet.Fonts = new Fonts();
            _stylesheet.Fonts.Count = 2;
            _stylesheet.Fonts.AppendChild(new Font());

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

            _stylesheet.Fonts.AppendChild(hyperLinkFont);

            _stylesheet.Borders = new Borders();
            _stylesheet.Borders.Count = 1;
            _stylesheet.Borders.AppendChild(new Border());

            _stylesheet.CellStyleFormats = new CellStyleFormats();
            _stylesheet.CellStyleFormats.Count = 2;
            _stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            _stylesheet.CellFormats = new CellFormats();
            // empty one for index 0, seems to be required
            _stylesheet.CellFormats.AppendChild(new CellFormat());
            CellFormat hyperLinkFormt = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            _stylesheet.CellFormats.AppendChild(hyperLinkFormt);
            _stylesheet.CellFormats.Count = 2;
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

        internal void AttachStylesPart(WorkbookPart workbookPart)
        {
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = _stylesheet;
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
                    newFont.FontName = new FontName() { Val = style.Font.FamilyName };
                    newFont.FontSize = new FontSize() { Val = style.Font.Size };
                    if (style.Font.Bold)
                    {
                        newFont.Bold = new Bold();
                    }
                    if (style.Font.Italic)
                    {
                        newFont.Italic = new Italic();
                    }
                    if (style.ForegroundColor != null)
                    {
                        newFont.Color = new Color() { Rgb = style.ForegroundColor.Value.GetRgbAsHex() };
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
                    newFill.ForegroundColor = new ForegroundColor() { Rgb = style.BackgroundColor.Value.GetRgbAsHex() };
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
            public FontKey(IronSoftware.Drawing.Font font, SixLabors.ImageSharp.Color? color)
            {
                Font = font;
                Color = color;
            }

            public IronSoftware.Drawing.Font Font { get; }
            public SixLabors.ImageSharp.Color? Color { get; }

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
