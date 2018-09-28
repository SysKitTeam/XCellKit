using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Formula = DocumentFormat.OpenXml.Office.Excel.Formula;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Acceleratio.XCellKit
{
    public abstract class SpreadsheetConditionalFormattingRule
    {

        public SpreadsheetLocation RangeStart { get; set; }
        public SpreadsheetLocation RangeEnd { get; set; }
        protected abstract ConditionalFormattingRuleTypeEnum Type { get; }
        public abstract void WriteOpenXml(OpenXmlWriter writer);

    }

    public class SpreadsheetConditionalSimpleIconSetFormattingRule : SpreadsheetConditionalFormattingRule
    {
        protected override ConditionalFormattingRuleTypeEnum Type
        {
            get
            {
                return ConditionalFormattingRuleTypeEnum.IconSet;
            }
        }

        public override void WriteOpenXml(OpenXmlWriter writer)
        {
           
            X14.ConditionalFormatting cf = new X14.ConditionalFormatting();            
            cf.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");
            ReferenceSequence referenceSequence = new ReferenceSequence();
            referenceSequence.Text = $"{SpreadsheetHelper.ExcelColumnFromNumber(RangeStart.ColumnIndex)}{RangeStart.RowIndex}:{SpreadsheetHelper.ExcelColumnFromNumber(RangeEnd.ColumnIndex)}{RangeEnd.RowIndex}";


            X14.ConditionalFormattingRule conditionalFormattingRule = new X14.ConditionalFormattingRule() { Type = ConditionalFormatValues.IconSet, Priority = 1, Id = Guid.NewGuid().ToString("B").ToUpper() };

            X14.IconSet iconSet1 = new X14.IconSet() { IconSetTypes = new EnumValue<IconSetTypeValues>(IconSetTypeValues.FourTrafficLights), ShowValue = false, Custom = true };
            //ConditionalFormattingIcon
            var conditionalFormattingIcon1 = new X14.ConditionalFormattingIcon() { IconSet = X14.IconSetTypeValues.NoIcons, IconId = (UInt32Value)0U };
            var conditionalFormattingIcon2 = new X14.ConditionalFormattingIcon() { IconSet = X14.IconSetTypeValues.ThreeSymbols2, IconId = (UInt32Value)2U };
            var conditionalFormattingIcon3 = new X14.ConditionalFormattingIcon() { IconSet = X14.IconSetTypeValues.ThreeSymbols, IconId = (UInt32Value)1U };
            var conditionalFormattingIcon4 = new X14.ConditionalFormattingIcon() { IconSet = X14.IconSetTypeValues.ThreeSymbols, IconId = (UInt32Value)0U };
            X14.ConditionalFormattingValueObject conditionalFormattingValueObject1 = new X14.ConditionalFormattingValueObject() { Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric };
            Formula formula1 = new Formula();
            formula1.Text = "0";
            conditionalFormattingValueObject1.Append(formula1);

            X14.ConditionalFormattingValueObject conditionalFormattingValueObject2 = new X14.ConditionalFormattingValueObject() { Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric };
            Formula formula2 = new Formula();
            formula2.Text = "1";
            conditionalFormattingValueObject2.Append(formula2);

            X14.ConditionalFormattingValueObject conditionalFormattingValueObject3 = new X14.ConditionalFormattingValueObject() { Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric };
            Formula formula3 = new Formula();
            formula3.Text = "2";
            conditionalFormattingValueObject3.Append(formula3);

            X14.ConditionalFormattingValueObject conditionalFormattingValueObject4 = new X14.ConditionalFormattingValueObject() { Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric };
            Formula formula4 = new Formula();
            formula4.Text = "3";
            conditionalFormattingValueObject4.Append(formula4);

            iconSet1.Append(
                  conditionalFormattingValueObject1
                , conditionalFormattingValueObject2
                , conditionalFormattingValueObject3
                , conditionalFormattingValueObject4
                , conditionalFormattingIcon1
                , conditionalFormattingIcon2
                , conditionalFormattingIcon3
                , conditionalFormattingIcon4
                );

            conditionalFormattingRule.Append(iconSet1);
            cf.Append(conditionalFormattingRule);
            cf.Append(referenceSequence);
            writer.WriteElement(cf);
        }
    }

    public enum ConditionalFormattingRuleTypeEnum
    {
        Expression,
        CellIs,
        ColorScale,
        DataBar,
        IconSet,
        Top10,
        UniqueValues,
        DuplicateValues,
        ContainsText,
        NotContainsText,
        BeginsWith,
        EndsWith,
        ContainsBlanks,
        NotContainsBlanks,
        ContainsErrors,
        NotContainsErrors,
        TimePeriod,
        AboveAverage,
    }
}
