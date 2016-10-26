using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace Acceleratio.XCellKit.Tests
{
    [TestClass]
    public class UnitTest2
    {
        [TestMethod]
        public void TestMethod1()
        {
            ConditionalFormattingRule rule = new ConditionalFormattingRule() { Type = ConditionalFormatValues.ColorScale, Priority = 1 };
            SpreadsheetConditionalFormatting format = new SpreadsheetConditionalFormatting("A1:A5", rule);
            Color first = new Color() { Rgb = "FFFFEF9C" };
            Color second = new Color() { Rgb = "EEEEEF9C" };
            format.formatBetweenMinAndMax(first, second);
            
        }
    }
}
