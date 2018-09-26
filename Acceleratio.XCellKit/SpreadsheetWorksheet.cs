using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Extension = DocumentFormat.OpenXml.Spreadsheet.Extension;
using ExtensionList = DocumentFormat.OpenXml.Spreadsheet.ExtensionList;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetWorksheet
    {
        const double DBL_AutoWidthCharacterScalingNumber = 1.1;
        public static int MaxUniqueHyperlinks = 20000;
        public string Name { get; set; }
        private int _maxColumnIndex;
        private int _maxRowIndex;

        private readonly Dictionary<int, int> _maxNumberOfCharsPerColumn = new Dictionary<int, int>();
        private readonly Dictionary<SpreadsheetLocation, SpreadsheetRow> _rows = new Dictionary<SpreadsheetLocation, SpreadsheetRow>();
        private readonly Dictionary<SpreadsheetLocation, SpreadsheetTable> _tables = new Dictionary<SpreadsheetLocation, SpreadsheetTable>();
        private readonly List<SpreadsheetConditionalFormattingRule> _conditionalFormattingRules = new List<SpreadsheetConditionalFormattingRule>();
        public DrawingsManager DrawingsManager { get; set; }
        public SpreadsheetWorksheet(string name)
        {
            Name = name;
            DrawingsManager = new DrawingsManager();
        }

        private bool _forzenFirstColumn = false;
        public void FreezeFirstColumn()
        {
            _forzenFirstColumn = true;
        }

        public void SetColumnWidth(int columnIndex, int charCount)
        {
            _maxNumberOfCharsPerColumn[columnIndex] = charCount;
        }

        public void AddTable(SpreadsheetTable table)
        {
            AddTable(table, 1, _maxRowIndex + 1);
        }

        public void AddTable(SpreadsheetTable table, int columnIndex, int rowIndex)
        {
            _tables[new SpreadsheetLocation(rowIndex, columnIndex)] = table;
            var headerRow = new SpreadsheetRow();
            for (var i = 0; i < table.Columns.Count; i++)
            {
                var column = table.Columns[i];
                var headerCell = new SpreadsheetCell();
                headerCell.Value = column.Name;
                headerRow.AddCell(headerCell);
                trackMaxChars(columnIndex + i, headerCell);
            }

            if (table.ShowHeaderRow)
            {
                addRow(headerRow, columnIndex, rowIndex, true);
                rowIndex++;
            }

            if (!table.IsInStreamingMode)
            {
                foreach (var row in table.Rows)
                {
                    AddRow(row, columnIndex, rowIndex);
                    rowIndex++;
                }
            }
            else
            {
                var enumerator = table.GetStreamingEnumerator();

                // nesto redaka cemo odmah dodati tako da se ispravno postave sirine stupaca
                // ostale retke cemo streamati kako se zapisuje u xlsx preko writera
                // MaxRowWidthsToTrackPerTable redaka nije puno za drzati u memoriji, ostatak ce ici 1 po 1 kak se pise sheet data
                var rowsToGet = MaxRowWidthsToTrackPerTable;
                var endOfTableIndex = rowIndex + table.RowCount;
                while (rowsToGet > 0 && enumerator.MoveNext())
                {
                    AddRow(enumerator.Current, columnIndex, rowIndex);
                    rowsToGet--;
                    rowIndex++;
                }
                if (endOfTableIndex > _maxRowIndex)
                {
                    _maxRowIndex = endOfTableIndex;
                }
            }
        }

        private void trackMaxChars(int columnIndex, SpreadsheetCell cell, bool isTableHeaderRow = false)
        {
            if (!cell.ParticipatesInAutoWidthColumnCalculation)
            {
                return;
            }
            var previousMax = 0;
            if (_maxNumberOfCharsPerColumn.ContainsKey(columnIndex))
            {
                previousMax = _maxNumberOfCharsPerColumn[columnIndex];
            }
            var charsCount = (cell.Value?.ToString().Split('\n').Max(x => x.Length) ?? 0) + cell.Indent;
            if (isTableHeaderRow)
            {
                charsCount += 4;
            }
            if (previousMax < charsCount)
            {
                _maxNumberOfCharsPerColumn[columnIndex] = charsCount;
            }
        }

        private void trackMaxChars(SpreadsheetRow row, SpreadsheetLocation location, bool isTableHeaderRow = false)
        {
            for (var i = 0; i < row.RowCells.Count; i++)
            {
                var cell = row.RowCells[i];
                trackMaxChars(location.ColumnIndex + i, cell, isTableHeaderRow);
            }
        }

        public void AddRow(SpreadsheetRow row)
        {
            AddRow(row, 1, _maxRowIndex + 1);
        }

        private int _rowWidthsTrackedSoFar = 0;
        const int MaxRowWidthsToTrack = 100000;
        private const int MaxRowWidthsToTrackPerTable = 5000;

        public void AddRow(SpreadsheetRow row, int columnIndex, int rowIndex)
        {
            addRow(row, columnIndex, rowIndex, false);
        }

        private void addRow(SpreadsheetRow row, int columnIndex, int rowIndex, bool isTableHeaderRow)
        {
            _rows[new SpreadsheetLocation(rowIndex, columnIndex)] = row;
            if (_rowWidthsTrackedSoFar < MaxRowWidthsToTrack)
            {
                trackMaxChars(row, new SpreadsheetLocation(rowIndex, columnIndex), isTableHeaderRow);
                _rowWidthsTrackedSoFar++;
            }
            var newMaxColumnIndex = columnIndex + row.RowCells.Count;
            if (newMaxColumnIndex > _maxColumnIndex)
            {
                _maxColumnIndex = newMaxColumnIndex;
            }
            if (rowIndex > _maxRowIndex)
            {
                _maxRowIndex = rowIndex;
            }
        }

        public void AddConditionalFormatting(SpreadsheetConditionalFormattingRule conditionalFormattingRule)
        {
            _conditionalFormattingRules.Add(conditionalFormattingRule);
        }

        public void WriteWorksheet(OpenXmlWriter writer, WorksheetPart part, SpreadsheetStylesManager stylesManager, ref int tableCount)
        {
            var hyperLinksManager = new SpreadsheetHyperlinkManager();
            writer.WriteStartElement(new Worksheet(), new List<OpenXmlAttribute>(), new List<KeyValuePair<string, string>>()
            {
               new KeyValuePair<string, string>( "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            });
            writeFrozenFirstColumn(writer);
            writeColumns(writer);
            writeSheetData(writer, stylesManager, hyperLinksManager, DrawingsManager);
            writeMergedCells(writer);
            writeHyperlinks(writer, part, hyperLinksManager);
            writeDrawings(part, writer);
            writeTables(writer, part, ref tableCount);
            writeExtensionsList(writer);
            writer.WriteEndElement();
        }

        private void writeDrawings(WorksheetPart part, OpenXmlWriter writer)
        {
            DrawingsManager.WriteDrawings(part, writer);
        }

        private void writeFrozenFirstColumn(OpenXmlWriter writer)
        {
            if (!_forzenFirstColumn)
            {
                return;
            }
            writer.WriteStartElement(new SheetViews());

            var tabSelectedAtt = new OpenXmlAttribute("tabSelected", null, 1.ToString());
            var workBookViewIdAtt = new OpenXmlAttribute("workbookViewId", null, 0.ToString());
            writer.WriteStartElement(new SheetView(), new List<OpenXmlAttribute>() { tabSelectedAtt, workBookViewIdAtt });

            var xSplitAtt = new OpenXmlAttribute("xSplit", null, 1.ToString());
            var topLeftCellAtt = new OpenXmlAttribute("topLeftCell", null, "B1");
            var activePane = new OpenXmlAttribute("activePane", null, "topRight");
            var state = new OpenXmlAttribute("state", null, "frozen");
            writer.WriteStartElement(new Pane(), new List<OpenXmlAttribute>() { xSplitAtt, topLeftCellAtt, activePane, state });
            writer.WriteEndElement();

            writer.WriteEndElement();

            writer.WriteEndElement();

        }

        private void writeHyperlinks(OpenXmlWriter writer, WorksheetPart woorksheetPart, SpreadsheetHyperlinkManager hyperlinkManager)
        {
            var hyperlinks = hyperlinkManager.GetHyperlinks();
            if (!hyperlinks.Any())
            {
                return;
            }

            var hyperlinkTargetRelationshipIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var clickableHyperLinks = new HashSet<KeyValuePair<SpreadsheetLocation, SpreadsheetHyperLink>>();

            // external ULRove moramo baciti kao relationshipove i onda ih u hyperlinkovima referencirat
            // zbog problema s trajanjem exporta u slucaju da ima puno hyperlinkova dodano je ogranicenje na INT_MaxUniqueHyperlinks, inace export traje 10x duze i to razlika moze samo rasti
            // jer je provjera za unique(za svaki link dodani) u pozadinskom openXML kodu sekvencijalan prolazak i usporedba kroz kolekciju. Odnosno O(n^2)
            // plus ovaj dio nije streamable nazalost, barem ne na ocit nacin
            foreach (var hyperlink in hyperlinks.Where(x => !string.IsNullOrEmpty(x.Value.Target)))
            {
                var target = hyperlink.Value.Target;
                if (hyperlinkTargetRelationshipIds.Count > MaxUniqueHyperlinks)
                {
                    break;
                }

                clickableHyperLinks.Add(hyperlink);
                if (!hyperlinkTargetRelationshipIds.ContainsKey(target))
                {
                    var uri = Utilities.SafelyCreateUri(target);
                    if (uri == null)
                    {
                        hyperlinkTargetRelationshipIds[target] = "";
                        continue;
                    }

                    var relId = woorksheetPart.AddHyperlinkRelationship(uri, true).Id;
                    hyperlinkTargetRelationshipIds[target] = relId;
                }
            }

            if (clickableHyperLinks.Count == 0)
            {
                return;
            }

            writer.WriteStartElement(new Hyperlinks());
            foreach (var link in clickableHyperLinks)
            {
                var attributes = new List<OpenXmlAttribute>();
                attributes.Add(new OpenXmlAttribute("ref", null, string.Format("{0}{1}", SpreadsheetHelper.ExcelColumnFromNumber(link.Key.ColumnIndex), link.Key.RowIndex)));
                string id;
                if (hyperlinkTargetRelationshipIds.TryGetValue(link.Value.Target, out id) && !string.IsNullOrEmpty(id))
                {
                    var idAtt = new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", hyperlinkTargetRelationshipIds[link.Value.Target]);
                    attributes.Add(idAtt);
                }
                else
                {
                    attributes.Add(new OpenXmlAttribute("location", null, link.Value.Target));
                    attributes.Add(new OpenXmlAttribute("display", null, link.Value.DisplayValue));
                }


                writer.WriteStartElement(new Hyperlink(), attributes);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private void writeSheetData(OpenXmlWriter writer, SpreadsheetStylesManager stylesManager, SpreadsheetHyperlinkManager hyperlinkManager, DrawingsManager drawingsManager)
        {
            writer.WriteStartElement(new SheetData());

            foreach (var row in _rows)
            {
                row.Value.WriteRow(writer, row.Key.ColumnIndex, row.Key.RowIndex, stylesManager, hyperlinkManager, drawingsManager);
            }

            foreach (var table in _tables)
            {
                if (!table.Value.IsInStreamingMode)
                {
                    continue;
                }
                var enumerator = table.Value.GetStreamingEnumerator();
                var tableRowPosition = table.Value.StreamedRowsSoFar;
                while (enumerator.MoveNext())
                {
                    var row = enumerator.Current;
                    row.WriteRow(writer, table.Key.ColumnIndex, table.Key.RowIndex + tableRowPosition + 1, stylesManager, hyperlinkManager, drawingsManager);
                    tableRowPosition++;
                }
            }

            writer.WriteEndElement();
        }

        private void writeMergedCells(OpenXmlWriter writer)
        {
            var mergedCellRanges = new Dictionary<SpreadsheetLocation, System.Drawing.Size>();
            foreach (var row in _rows)
            {
                for (var i = 0; i < row.Value.RowCells.Count; i++)
                {
                    var cell = row.Value.RowCells[i];
                    if (cell.MergedCellsRange != null)
                    {
                        mergedCellRanges[new SpreadsheetLocation(row.Key.RowIndex, row.Key.ColumnIndex + i)] = cell.MergedCellsRange.Value;
                    }
                }
            }

            if (mergedCellRanges.Any())
            {
                writer.WriteStartElement(new MergeCells());

                foreach (var cellRange in mergedCellRanges)
                {
                    var cell1Name = SpreadsheetHelper.ExcelColumnFromNumber(cellRange.Key.ColumnIndex) + cellRange.Key.RowIndex;
                    var cell2Name = SpreadsheetHelper.ExcelColumnFromNumber(cellRange.Key.ColumnIndex + cellRange.Value.Width) + (cellRange.Key.RowIndex + cellRange.Value.Height - 1);
                    var range = cell1Name + ":" + cell2Name;
                    var mergeCell = new MergeCell()
                    {

                        Reference = new StringValue(range)
                    };
                    writer.WriteElement(mergeCell);
                }
                writer.WriteEndElement();
            }
        }

        private void writeExtensionsList(OpenXmlWriter writer)
        {
            if (_conditionalFormattingRules.Count == 0)
            {
                return;
            }
            writer.WriteStartElement(new ExtensionList());
            writer.WriteStartElement(new Extension(), new List<OpenXmlAttribute>()
            {
                new OpenXmlAttribute("uri", null, "{78C0D931-6437-407d-A8EE-F0AAD7539E65}")
            },
            new List<KeyValuePair<string, string>>()
            {
               new KeyValuePair<string, string>( "x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
            });
            writeConditionalFormattings(writer);
            writer.WriteEndElement();
        }

        private void writeConditionalFormattings(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new X14.ConditionalFormattings());
            foreach (var spreadsheetConditionalFormattingRule in _conditionalFormattingRules)
            {
                spreadsheetConditionalFormattingRule.WriteOpenXml(writer);
            }
            writer.WriteEndElement();
        }


        private void writeTables(OpenXmlWriter writer, WorksheetPart part, ref int tableCount)
        {
            if (!_tables.Any())
            {
                return;
            }
            var countAtt = new OpenXmlAttribute("count", null, _tables.Count.ToString());
            writer.WriteStartElement(new TableParts(), new List<OpenXmlAttribute>() { countAtt });

            foreach (var table in _tables)
            {
                var tableId = "table" + tableCount;
                var tableDefinition = part.AddNewPart<TableDefinitionPart>(tableId);
                tableDefinition.Table = table.Value.GetTableDefinition(tableCount, table.Key.ColumnIndex, table.Key.RowIndex);
                var idAtt = new OpenXmlAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", tableId);
                writer.WriteStartElement(new TablePart(), new List<OpenXmlAttribute>() { idAtt });
                writer.WriteEndElement();
                tableCount++;
            }
            writer.WriteEndElement();
        }

        private void writeColumns(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Columns());
            for (var i = 1; i <= _maxColumnIndex; i++)
            {
                var width = 20d;
                if (_maxNumberOfCharsPerColumn.ContainsKey(i))
                {
                    width = _maxNumberOfCharsPerColumn[i] > 255 ? 255 : _maxNumberOfCharsPerColumn[i];
                    width = width * DBL_AutoWidthCharacterScalingNumber;
                }
                var minAtt = new OpenXmlAttribute("min", null, i.ToString());
                var maxAtt = new OpenXmlAttribute("max", null, i.ToString());

                if (MaxColumnWidth.HasValue && width > MaxColumnWidth)
                {
                    width = MaxColumnWidth.Value;
                }

                var widthAtt = new OpenXmlAttribute("width", null, width.ToString(CultureInfo.InvariantCulture));
                writer.WriteStartElement(new Column(), new List<OpenXmlAttribute>() { minAtt, maxAtt, widthAtt });
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public int LastRowIndex
        {
            get { return _maxRowIndex; }
        }

        public double? MaxColumnWidth { get; set; } = null;
    }
}
