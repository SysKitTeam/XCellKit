using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acceleratio.XCellKit
{
    class SpreadsheetPivotTableDefinition
    {
        private Location location;
        private PivotFields fields;
        private RowFields rowFields;
        private RowItems rowItems;
        private ColumnItems columnItems;
        private DataFields dataFields;
        private PivotTableDefinition definition;
        private List<PivotField> listPivotField;
        private List<Items> listItems;
        private List<Item> listItem;
        private List<Field> listField;
        private List<RowItem> listRowItem;
        private List<DataField> listDataField;
        private PivotTableStyle tableStyle;


        public SpreadsheetPivotTableDefinition(string name, int cacheID)
        {
            definition = new PivotTableDefinition() { Name = name, CacheId = (UInt32Value)(UInt32)cacheID };
        }

        public void Location(string reference, int firstHeaderRow, int firstDataRow, int firstDataColumn)
        {
            location = new Location()
            {
                Reference = reference,
                FirstHeaderRow = (UInt32Value)(UInt32)firstHeaderRow,
                FirstDataRow = (UInt32Value)(UInt32)firstDataRow,
                FirstDataColumn = (UInt32Value)(UInt32)firstDataColumn
            };
        }

        //Add item with index value to listItem
        public void AddItem(int index)
        {
            Item item = new Item() { Index = (UInt32Value)(UInt32)index };
            listItem.Add(item);
        }

        /*
        Initialize Items variable and appends all of the item in the list to items,
        add items variable to listItems and clears the listItem
        */
        public void ListItemToItems()
        {
            Items items = new Items() { Count = (UInt32Value)(UInt32)listItem.Count };
            foreach(var item in listItem)
            {
                items.Append(item);
            }
            listItems.Add(items);
            listItem.Clear();
        }

        /*
        Initialize PivotField variable and appends all of the items in the list to field,
        add field variable to listPivotField and clears the listItems
        */
        public void ListItemsToPivotField()
        {
            PivotField field = new PivotField();
            foreach (var items in listItems)
            {
                field.Append(items);
            }
            listPivotField.Add(field);
            listItems.Clear();
        }

        /*
        Initialize PivotFields variable and appends all of the pivot field in the list to pivot fields
        and clears the listPivotField
        */
        public void ListPivotFieldToPivotFields()
        {
            fields = new PivotFields() { Count = (UInt32Value)(UInt32)listPivotField.Count };
            foreach (var field in listPivotField)
            {
                fields.Append(field);
            }
            listPivotField.Clear();
        }

        //Add field with index value to listField
        public void AddField(int index)
        {
            Field field = new Field() { Index = index };
            listField.Add(field);
        }

        /*
        Initialize RowFields variable and appends all of the fields in the list to row fields
        and clears the listField
        */
        public void ListFieldToRowFields()
        {
            rowFields = new RowFields() { Count = (UInt32Value)(UInt32)listField.Count };
            foreach (var field in listField)
            {
                rowFields.Append(field);
            }
            listField.Clear();
        }

        /*
        Initialize RowItem variable, if the parametar containValue is true then the method will
        take the val parametar and initialize new MemberProperyIndex variable and set it's Val to
        parametar val, else it just initialize new MemberProperyIndex variable
        */
        public void AddRowItem(bool containValue, int val=0)
        {
            RowItem rowItem = new RowItem();
            MemberPropertyIndex member;
            if (containValue)
            {
                member = new MemberPropertyIndex() { Val = val };
            }
            else
            {
                member = new MemberPropertyIndex();
            }
            rowItem.Append(member);
            listRowItem.Add(rowItem);
        }

        /*
        Initialize RowItems variable and appends all of the row item in the list to row items
        and clears the listRowItem
        */
        public void ListRowItemToRowItems()
        {
            rowItems = new RowItems() { Count = (UInt32Value)(UInt32)listRowItem.Count };
            foreach (var item in listRowItem)
            {
                rowItems.Append(item);
            }
            listRowItem.Clear();
        }

        /*
        Initialize ColumnItems variable and appends all of the row item in the list to column items
        and clears the listRowItem
        */
        public void ListRowItemToColumnItems()
        {
            columnItems = new ColumnItems() { Count = (UInt32Value)(UInt32)listRowItem.Count };
            foreach (var item in listRowItem)
            {
                columnItems.Append(item);
            }
            listRowItem.Clear();
        }

        //Initialize DataField variable and add it to listDataField
        public void AddDataField(string name, int field, int baseField, int baseItem)
        {
            DataField dataField = new DataField()
            {
                Name = name,
                Field = (UInt32Value)(UInt32)field,
                BaseField = baseField,
                BaseItem = (UInt32Value)(UInt32)baseItem
            };
            listDataField.Add(dataField);
        }

        /*
        Initialize DataFields variable and appends all of the data fields in the list to data fields
        and clears the listRowItem
        */
        public void ListDataFieldToDataFields()
        {
            dataFields = new DataFields() { Count = (UInt32Value)(UInt32)listDataField.Count };
            foreach (var item in listDataField)
            {
                dataFields.Append(item);
            }
            listDataField.Clear();
        }

        //Initialize PivotTableStyle variable and sets some of the initial values
        public void Style(string name, bool showRowHeaders, bool showColumnHeaders, bool showRowStripes, bool showColumnStripes, bool showLastColumn)
        {
            tableStyle = new PivotTableStyle()
            {
                Name = name,
                ShowRowHeaders = showRowHeaders,
                ShowColumnHeaders = showColumnHeaders,
                ShowColumnStripes = showColumnStripes,
                ShowRowStripes = showRowStripes,
                ShowLastColumn = showLastColumn
            };
        }

        //Appends all the elements to the pivot table definition and returns it
        public PivotTableDefinition GeneratePivotTableDefinition()
        {
            definition.Append(location);
            definition.Append(fields);
            definition.Append(rowFields);
            definition.Append(rowItems);
            definition.Append(columnItems);
            definition.Append(dataFields);
            definition.Append(tableStyle);

            return definition;
        }
    }
}
