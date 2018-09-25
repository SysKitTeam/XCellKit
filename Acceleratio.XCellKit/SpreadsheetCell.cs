using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public enum SpreadsheetDataTypeEnum
    {
        String,
        Number,
        DateTime,
        Other
    }
    public class SpreadsheetCell
    {
        public object Value { get; set; }
        public System.Drawing.Font Font { get; set; }
        public System.Drawing.Color? BackgroundColor { get; set; }
        public System.Drawing.Color? ForegroundColor { get; set; }
        public HorizontalAligment? Alignment { get; set; }
        public int Indent { get; set; }
        public SpreadsheetDataTypeEnum SpreadsheetDataType { get; set; }

        public SpreadsheetCell()
        {
            Indent = 0;
            SpreadsheetDataType = SpreadsheetDataTypeEnum.String;
        }

        public virtual void WriteCell(OpenXmlWriter writer, int columnIndex, int rowIndex, SpreadsheetStylesManager stylesManager, SpreadsheetHyperlinkManager hyperlinkManager)
        {
            var openXmlAtts = new List<OpenXmlAttribute>();
            var columnLetter = SpreadsheetHelper.ExcelColumnFromNumber(columnIndex);
            var position = string.Format("{0}{1}", columnLetter, rowIndex);
            var positionAtt = new OpenXmlAttribute("r", null, position);
            openXmlAtts.Add(positionAtt);
            
            var styleAtt = getStyleAttribute(stylesManager);
            if (styleAtt.HasValue)
            {
                openXmlAtts.Add(styleAtt.Value);
            }

            var sValue = Value.ToString();
            
            // Total number of characters that a cell can contain is 32,767.
            if (sValue.Length > 32767)
            {
                sValue = sValue.Substring(0, 32767);
            }
            if (SpreadsheetDataType == SpreadsheetDataTypeEnum.Number)
            {
                double numberValue = 0;
                if (double.TryParse(sValue, out numberValue))
                {
                    var typeAtt = new OpenXmlAttribute("t", null, "n");
                    openXmlAtts.Add(typeAtt);
                    writer.WriteStartElement(new Cell(), openXmlAtts);
                    writer.WriteElement(new CellValue(numberValue.ToString(CultureInfo.InvariantCulture)));
                    writer.WriteEndElement();
                }
            }
            else if (SpreadsheetDataType == SpreadsheetDataTypeEnum.String)
            {
                var typeAtt = new OpenXmlAttribute("t", null, "inlineStr");
                openXmlAtts.Add(typeAtt);
                writer.WriteStartElement(new Cell(), openXmlAtts);
                sValue = XmlConvert.EncodeName(sValue);
                writer.WriteStartElement(new InlineString());
                writer.WriteElement(new Text(sValue) { Space = SpaceProcessingModeValues.Preserve  });
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            else if (SpreadsheetDataType == SpreadsheetDataTypeEnum.DateTime)
            {
                var dateTime = DateTime.MinValue;
                if (DateTime.TryParse(sValue, out dateTime))
                {
                    writer.WriteStartElement(new Cell(), openXmlAtts);
                    writer.WriteElement(new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture)));
                    writer.WriteEndElement();
                }
            }
            else if (SpreadsheetDataType == SpreadsheetDataTypeEnum.Other)
            {
                var typeAttribute = new OpenXmlAttribute("t", null, "str");
                openXmlAtts.Add(typeAttribute);
                writer.WriteStartElement(new Cell(), openXmlAtts);
                writer.WriteElement(new CellValue(sValue));
                writer.WriteEndElement();
            }
           
            
            
        }

        protected virtual OpenXmlAttribute? getStyleAttribute(SpreadsheetStylesManager stylesManager)
        {
            OpenXmlAttribute? styleAtt = null;
            if (Font != null || BackgroundColor != null || ForegroundColor != null || Alignment != null || Indent != 0 || SpreadsheetDataType == SpreadsheetDataTypeEnum.DateTime)
            {
                var spreadsheetStyle = new SpreadsheetStyle()
                {
                    Font = Font,
                    BackgroundColor = BackgroundColor,
                    ForegroundColor = ForegroundColor,
                    Alignment = Alignment.HasValue ? SpreadsheetHelper.GetHorizontalAlignmentValue(Alignment.Value) : (HorizontalAlignmentValues?) null,
                    Indent = Indent,
                    IsDate =  SpreadsheetDataType == SpreadsheetDataTypeEnum.DateTime
                };
                styleAtt = new OpenXmlAttribute("s", null, ((UInt32)stylesManager.GetStyleIndex(spreadsheetStyle)).ToString());
            }
            return styleAtt;
        }
    }
}
