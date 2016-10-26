using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetConditionalFormatting
    {
        private ConditionalFormatting conditionalFormatting;
        private ConditionalFormattingRule conditionalRule;

        public SpreadsheetConditionalFormatting(string reference, ConditionalFormattingRule rule)
        {
            conditionalRule = rule;
            conditionalFormatting = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = reference } };
        }

        public ConditionalFormatting formatBetweenMinAndMax(Color first, Color second)
        {
            ColorScale colorScale = new ColorScale();
            ConditionalFormatValueObject conditionalFormatValueObject1 = new ConditionalFormatValueObject() { Type = ConditionalFormatValueObjectValues.Min};
            ConditionalFormatValueObject conditionalFormatValueObject2 = new ConditionalFormatValueObject() { Type = ConditionalFormatValueObjectValues.Max };
            colorScale.Append(conditionalFormatValueObject1);
            colorScale.Append(conditionalFormatValueObject2);
            colorScale.Append(first);
            colorScale.Append(second);
            conditionalRule.Append(colorScale);
            conditionalFormatting.Append(conditionalRule);
            return conditionalFormatting;
        }

        public void writeConditionalFormats(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new ConditionalFormatting());
            writer.WriteElement(conditionalFormatting);
            writer.WriteEndElement();
        }


    }
}
