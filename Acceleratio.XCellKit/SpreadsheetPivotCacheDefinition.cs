using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Acceleratio.XCellKit
{
    class SpreadsheetPivotCacheDefinition
    {
        private List<StringItem> listStringItem;
        public static PivotCacheDefinition cacheDefinition;
        private CacheField field;
        private List<CacheField> listField;
        private CacheFields cacheFields;
        private List<PivotCacheDefinitionExtension> extensions;
        private CacheSource cache;

        /*Constructor that initialize the main PivotCacheDefinition element,
          sets the id and record count and initialize CacheFields variable*/
        public SpreadsheetPivotCacheDefinition(string id)
        {
            cacheDefinition = new PivotCacheDefinition()
            {
                Id = id,
                RecordCount = SpredsheetPivotCacheRecords.records.Count
            };
            cacheDefinition.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            CacheFields cacheFields = new CacheFields();
        }

        //TODO: pronadi ostale moguce source-ove
        public void WorksheetSource(string source, string sheet, string id)
        {
            WorksheetSource sheetSource = new WorksheetSource() { Reference = source, Id = id };
            sheetSource.Sheet = sheet;
            cache = new CacheSource() { Type = SourceValues.Worksheet };
            cache.Append(sheetSource);
        }

        //initialize new cache field with name and a number format
        public void Field(string name, int numberFormat)
        {
            field = new CacheField()
            {
                Name = name,
                NumberFormatId = (UInt32Value)(UInt32)numberFormat
            };
            listField.Add(field);
        }

        //create new string item and add it to the list
        public void NewStringItem(string val)
        {
            StringItem item = new StringItem() { Val = val };
            listStringItem.Add(item);
        }

        /*Create new SharedItems element, add all elements from listStringItem, 
          add SharedItems to CacheField and clear the string item list*/
        public void AddStringItemsToSharedItems()
        {
            SharedItems sharedItems = new SharedItems()
            {
                Count = (UInt32Value)(UInt32)listStringItem.Count
            };
            foreach (var item in listStringItem)
            {
                sharedItems.Append(item);
            }
            field.Append(sharedItems);
            listStringItem.Clear();
        }

        //Create new shared item for integer numbers with min and max value
        public void IntegerItems(int min, int max)
        {
            SharedItems sharedItems = new SharedItems()
            {
                ContainsSemiMixedTypes = false,
                ContainsString = false,
                ContainsNumber = true, ContainsInteger = true,
                MinValue = (DoubleValue)min,
                MaxValue = (DoubleValue)max
            };
            field.Append(sharedItems);
        }

        //add all cache fields form list to the element CacheFields
        public void AddCacheFields()
        {
            if(listField.Count == 0)
            {
                System.Exception e = new System.Exception("Number of generated Cache fields is 0");
                throw e;
            }
            else
            {
                foreach (var item in listField)
                {
                    cacheFields.Append(item);
                }
            }
        }

        public void GeneratePivotCache()
        {
            cacheDefinition.Append(cache);
            cacheDefinition.Append(cacheFields);
        }
    }
}
