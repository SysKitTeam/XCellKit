using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SysKit.XCellKit
{
    public partial class SpreadsheetTable
    {
        /// <summary>
        /// Used for dynamic row building when in streaming mode
        /// If an enumerator was not provided while calling <see cref="ActivateStreamingMode"/>
        /// </summary>
        public event EventHandler<RequestTableRowEventArgs> TableRowRequested = delegate { };

        private void RaiseRequestTableRow(RequestTableRowEventArgs args)
        {
            if (!IsInStreamingMode)
            {
                throw new InvalidOperationException("Cannot request rows in non streaming mode");
            }
            TableRowRequested(this, args);
        }
        private List<SpreadsheetRow> _rows;
        private bool _streamingMode;
        private SpreadSheetTableStreamingEnumerator _streamingEnumerator;
        public SpreadsheetTable(string name)
        {
            Name = name;
            Columns = new List<SpreadsheetTableColumn>();
            _rows = new List<SpreadsheetRow>();
            _streamingMode = false;
            FillLastCellInRow = true;
            ShowHeaderRow = true;
        }
        public string Name { get; private set; }
        public List<SpreadsheetTableColumn> Columns { get; set; }

        public int RowCount
        {
            get
            {
                if (_streamingMode)
                {

                    if (_streamingEnumerator.ExhaustedAllRows)
                    {
                        return _streamingEnumerator.ItemsRead;
                    }
                    throw new InvalidOperationException("Row count is not available in streaming mode when not all rows have been read");
                }
                return Rows.Count;
            }
        }

        public List<SpreadsheetRow> Rows
        {
            get
            {
                if (_streamingMode)
                {
                    throw new InvalidOperationException("Direct row access is not allowed in streaming mode");
                }
                return _rows;
            }
            set
            {
                if (_streamingMode)
                {
                    throw new InvalidOperationException("Direct row access is not allowed in streaming mode");
                }
                _rows = value;
            }
        }

        /// <summary>
        /// Calling this function will activate the streaming mode when exporting the table.
        /// Rows can be provided by using the <see cref="TableRowRequested"/> event.
        /// </summary>
        public void ActivateStreamingMode()
        {
            ActivateStreamingMode(null);
        }

        /// <summary>
        /// Calling this function will activate the streaming mode when exporting the table.
        /// Rows will be read by using the <paramref name="rowsEnumerator"/> parameter.
        /// <para>
        /// To get the most out of the streaming functionality please plan ahead about the way you provide the data to the table.
        /// </para>
        /// Any data manipulations should occur without additional .ToList calls or additional lists.
        /// Ie. use .Select and transform your data to rows without calling .ToList, pass the enumerator of the resulting IEnumerable.
        /// </summary>     
        public void ActivateStreamingMode(IEnumerator<SpreadsheetRow> rowsEnumerator)
        {
            if (IsInStreamingMode && this.StreamedRowsSoFar > 0)
            {
                throw new InvalidOperationException("Streaming mode cannot be activated once rows have been added");
            }
            _streamingMode = true;

            _streamingEnumerator = new SpreadSheetTableStreamingEnumerator(this, rowsEnumerator);
        }

      
        public int StreamedRowsSoFar
        {
            get
            {
                if (!_streamingMode)
                {
                    throw new InvalidOperationException("Cannot read while not in streaming mode");
                }
                return _streamingEnumerator.ItemsRead;
            }
        }

        public IEnumerator<SpreadsheetRow> GetStreamingEnumerator()
        {
            return _streamingEnumerator;
        }

        public bool IsInStreamingMode
        {
            get { return _streamingMode; }
        }

        public bool FillLastCellInRow { get; set; }

        public bool ShowHeaderRow { get; set; }

        public Table GetTableDefinition(int id, int columnIndex, int rowIndex)
        {
            var startColumn = SpreadsheetHelper.ExcelColumnFromNumber(columnIndex);
            var endColumn = SpreadsheetHelper.ExcelColumnFromNumber(columnIndex + Columns.Count - 1);

            var reference = string.Format("{0}{1}:{2}{3}", startColumn, rowIndex, endColumn, rowIndex + RowCount + (ShowHeaderRow ? 0 : -1));

            var table = new Table() { Id = (UInt32Value)(UInt32)id, Name = Name, DisplayName = "table" + id + "_", TotalsRowShown = false, Reference = reference };
            AutoFilter autoFilter = null;

            if (!ShowHeaderRow)
            {
                table.HeaderRowCount = 0;
            }
            else
            {
                autoFilter = new AutoFilter() { Reference = reference };
            }
            var tablesColumns = new TableColumns() { Count = (UInt32Value)(UInt32)Columns.Count };
            var i = 1;
            foreach (var column in Columns)
            {
                tablesColumns.Append(column.ToOpenXmlColumn(i));
                if (autoFilter != null && column.FilterValues.Any())
                {
                    var filterColumn = new FilterColumn() { ColumnId = (UInt32)i - 1 };
                    var filters = new Filters();
                    foreach (var filterValue in column.FilterValues)
                    {
                        filters.AppendChild(new Filter() { Val = filterValue });
                    }
                    filterColumn.AppendChild(filters);

                    autoFilter.AppendChild(filterColumn);
                }
                i++;
            }
            TableStyleInfo tableStyle = new TableStyleInfo() { Name = "TableStyleLight9", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false, };
            if (autoFilter != null)
            {
                table.Append(autoFilter);
            }
            table.Append(tablesColumns);
            table.Append(tableStyle);

            return table;

        }
    }
}
