using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SysKit.XCellKit
{
    public class SpreadsheetHyperlinkCell : SpreadsheetCell
    {
        private readonly SpreadsheetHyperLink _hyperLink;
        public SpreadsheetHyperlinkCell(SpreadsheetHyperLink hyperLink)
        {
            _hyperLink = hyperLink;
            Value = hyperLink.DisplayValue;
        }

        protected override OpenXmlAttribute? getStyleAttribute(SpreadsheetStylesManager stylesManager)
        {
            var spreadsheetStyle = new SpreadsheetStyle();
            if (Indent != 0 || Alignment != null || VerticalAlignment != null)
            {
                spreadsheetStyle = new SpreadsheetStyle()
                {
                    Alignment = Alignment.HasValue ? SpreadsheetHelper.GetHorizontalAlignmentValue(Alignment.Value) : (HorizontalAlignmentValues?)null,
                    VerticalAlignment = VerticalAlignment.HasValue ? SpreadsheetHelper.GetVerticalAlignmentValues(VerticalAlignment.Value) : (VerticalAlignmentValues?)null,
                    BackgroundColor = BackgroundColor,
                    Font = Font,
                    ForegroundColor = ForegroundColor,
                    Indent = Indent
                };
            }
            return new OpenXmlAttribute("s", null, ((UInt32)stylesManager.GetHyperlinkStyleIndex(spreadsheetStyle)).ToString());
        }

        public override void WriteCell(OpenXmlWriter writer, int columnIndex, int rowIndex, SpreadsheetStylesManager stylesManager, SpreadsheetHyperlinkManager hyperlinkManager)
        {
            base.WriteCell(writer, columnIndex, rowIndex, stylesManager, hyperlinkManager);
            hyperlinkManager.AddHyperlink(new SpreadsheetLocation(rowIndex, columnIndex), _hyperLink);
        }
    }
}
