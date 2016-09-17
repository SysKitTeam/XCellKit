using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    class SpredsheetPivotCacheRecords
    {
        public static PivotCacheRecords records;
        private List<PivotCacheRecord> listRecord;

        public SpredsheetPivotCacheRecords()
        {
            records = new PivotCacheRecords();
            records.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        }

        public NumberItem AddNumberItem(int val)
        {
            NumberItem item = new NumberItem() { Val = (DoubleValue)val };
            return item;
        }

        public FieldItem AddFieldItem(int val)
        {
            FieldItem item = new FieldItem() { Val = (UInt32Value)(UInt32)val };
            return item;
        }

        public void GenerateRecord()
        {
            PivotCacheRecord record = new PivotCacheRecord();
            listRecord.Add(record);
        }

        public void AppendNumberToCacheRecod(NumberItem item)
        {
            listRecord.Last().Append(item);
        }

        public void AppendFieldToCacheRecod(FieldItem item)
        {
            listRecord.Last().Append(item);
        }

        public void GeneratePivotCacheRecords()
        {
            records.Count = (UInt32Value)(UInt32)listRecord.Count;
            foreach (var item in listRecord)
            {
                records.Append(item);
            }
        }


    }
}
