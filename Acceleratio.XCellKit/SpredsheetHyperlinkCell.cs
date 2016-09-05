using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class SpredsheetHyperlinkCell : SpredsheetCell
    {
        private SpredsheetHyperLink _hyperLink;
        public SpredsheetHyperlinkCell(SpredsheetHyperLink hyperLink)
        {
            _hyperLink = hyperLink;
        }

        protected override OpenXmlAttribute? getStyleAttribute(SpredsheetStylesManager stylesManager)
        {
            var spredsheetStyle = new SpredsheetStyle();
            if (Indent != 0 || Alignment != null)
            {
                spredsheetStyle = new SpredsheetStyle()
                {
                    Alignment = Alignment.HasValue ? SpredsheetHelper.GetHorizontalAlignmentValue(Alignment.Value) : (HorizontalAlignmentValues?)null,
                    BackgroundColor = BackgroundColor,
                    Font = Font,
                    ForegroundColor = ForegroundColor,
                    Indent = Indent
                };
            }
            return new OpenXmlAttribute("s", null, ((UInt32)stylesManager.GetHyperlinkStyleIndex(spredsheetStyle)).ToString());
        }

        public override void WriteCell(OpenXmlWriter writer, int columnIndex, int rowIndex, SpredsheetStylesManager stylesManager, SpredsheetHyperlinkManager hyperlinkManager)
        {
            base.WriteCell(writer, columnIndex, rowIndex, stylesManager, hyperlinkManager);
            hyperlinkManager.AddHyperlink(new SpredsheetLocation(rowIndex, columnIndex), _hyperLink);
        }
    }
}
