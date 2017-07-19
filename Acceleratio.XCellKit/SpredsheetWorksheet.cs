using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class SpredsheetWorksheet
    {
        public string Name { get; set; }
        private int _maxColumnIndex;
        private int _maxRowIndex;
        private Dictionary<int, int> _maxNumberOfCharsPerColumn; 

        public SpredsheetWorksheet(string name)
        {
            Name = name;
            _tables = new Dictionary<SpredsheetLocation, SpredsheetTable>();
            _rows = new Dictionary<SpredsheetLocation, SpredsheetRow>();
            _maxNumberOfCharsPerColumn = new Dictionary<int, int>();
            _charts = new Dictionary<SpredsheetLocation, SpredsheetChart>();
        }

        public SpredsheetWorksheet(string name, List<int> columnsIndexToTrackAutoWidht)
            :this(name)
        {
            foreach (var i in columnsIndexToTrackAutoWidht)
            {
                _maxNumberOfCharsPerColumn[i] = 0;
            }
        }

        private bool _forzenFirstColumn = false;
        public void FreezeFirstColumn()
        {
            _forzenFirstColumn = true;
        }

        private Dictionary<SpredsheetLocation, SpredsheetTable> _tables; 
        public void AddTable(SpredsheetTable table, int columnIndex, int rowIndex)
        {
            _tables[new SpredsheetLocation(rowIndex, columnIndex)] = table;
            var headerRow = new SpredsheetRow();
            for (int i = 0; i<table.Columns.Count; i++)
            {
                var column = table.Columns[i];
                var headerCell = new SpredsheetCell();
                headerCell.Value = column.Name;
                headerRow.AddCell(headerCell);
                if (_maxNumberOfCharsPerColumn.Any())
                {
                    trackMaxChars(columnIndex + i, headerCell);
                }
            }

            AddRow(headerRow, columnIndex, rowIndex);
            rowIndex++;
            foreach (var row in table.Rows)
            {
                AddRow(row, columnIndex, rowIndex);
                rowIndex++;
            }
        }

        private Dictionary<SpredsheetLocation, SpredsheetChart> _charts;
        public void AddChart(SpredsheetChart chart, int columnIndex, int rowIndex)
        {
            _charts[new SpredsheetLocation(rowIndex, columnIndex)] = chart;
        }

        private void trackMaxChars(int columnIndex, SpredsheetCell cell)
        {
            var previousMax = _maxNumberOfCharsPerColumn[columnIndex];
            var charsCount = cell.Value.ToString().Count() + cell.Indent; 
            if (previousMax < charsCount)
            {
                _maxNumberOfCharsPerColumn[columnIndex] = charsCount;
            }
        }

        private void trackMaxChars(SpredsheetRow row, SpredsheetLocation location)
        {
            for (int i = 0; i < row.RowCells.Count; i++)
            {
                var cell = row.RowCells[i];
                trackMaxChars(location.ColumnIndex + i, cell);
            }
        }
        
        public void AddRow(SpredsheetRow row)
        {
            AddRow(row, 1, _maxRowIndex + 1);
        }

        private Dictionary<SpredsheetLocation, SpredsheetRow> _rows;  
        public void AddRow(SpredsheetRow row, int columnIndex, int rowIndex)
        {
            _rows[new SpredsheetLocation(rowIndex, columnIndex)] = row;
            if (_maxNumberOfCharsPerColumn.Any())
            {
                trackMaxChars(row, new SpredsheetLocation(rowIndex, columnIndex));
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
        
        public void WriteWorksheet(OpenXmlWriter writer, WorksheetPart part, SpredsheetStylesManager stylesManager, ref int tableCount)
        {
            var hyperLinksManager = new SpredsheetHyperlinkManager();
            var newWorksheet = new Worksheet();
            writer.WriteStartElement(newWorksheet);
            writeFrozenFirstColumn(writer);
            writeColumns(writer);
            writeSheetData(writer, stylesManager, hyperLinksManager);
            writeHyperlinks(writer, hyperLinksManager);
            writeChart(writer, part);
            writeTables(writer, part, ref tableCount);
            writer.WriteEndElement();
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
            writer.WriteStartElement(new SheetView(), new List<OpenXmlAttribute>() {tabSelectedAtt, workBookViewIdAtt});

            var xSplitAtt = new OpenXmlAttribute("xSplit", null, 1.ToString());
            var topLeftCellAtt = new OpenXmlAttribute("topLeftCell", null, "B1");
            var activePane = new OpenXmlAttribute("activePane", null, "topRight");
            var state = new OpenXmlAttribute("state", null, "frozen");
            writer.WriteStartElement(new Pane(), new List<OpenXmlAttribute>() { xSplitAtt, topLeftCellAtt, activePane, state});
            writer.WriteEndElement();
            
            writer.WriteEndElement();

            writer.WriteEndElement();

        }
        
        private void writeHyperlinks(OpenXmlWriter writer, SpredsheetHyperlinkManager hyperlinkManager)
        {
            var hyperlinks = hyperlinkManager.GetHyperlinks();
            if (!hyperlinks.Any())
            {
                return;
            }
            writer.WriteStartElement(new Hyperlinks());
            foreach (var link in hyperlinks)
            {
                link.Value.WriteHyperLink(writer, link.Key.ColumnIndex, link.Key.RowIndex);
            }
            writer.WriteEndElement();
        }

        private void writeSheetData(OpenXmlWriter writer, SpredsheetStylesManager stylesManager, SpredsheetHyperlinkManager hyperlinkManager)
        {
            writer.WriteStartElement(new SheetData());

            foreach (var row in _rows)
            {
                row.Value.WriteRow(writer, row.Key.ColumnIndex, row.Key.RowIndex, stylesManager, hyperlinkManager);
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
            writer.WriteStartElement(new TableParts(), new List<OpenXmlAttribute>() { countAtt});

            foreach (var table in _tables)
            {
                var tableId = "table" + tableCount;
                var tableDefinition = part.AddNewPart<TableDefinitionPart>(tableId);
                tableDefinition.Table = table.Value.GetTableDefinition(tableCount, table.Key.ColumnIndex, table.Key.RowIndex);
               
                var idAtt = new OpenXmlAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", tableId);
                writer.WriteStartElement(new TablePart(), new List<OpenXmlAttribute>() { idAtt});
                writer.WriteEndElement();
                tableCount++;
            }
            writer.WriteEndElement();
        }

        public void writeChart(OpenXmlWriter writer, WorksheetPart part)
        {
            foreach (var singleChart in _charts)
            {
                DrawingsPart drawingsPart = part.AddNewPart<DrawingsPart>();

                writer.WriteStartElement(new Drawing() { Id = part.GetIdOfPart(drawingsPart) });
                writer.WriteEndElement();

                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });

                singleChart.Value.CreateChart(chartPart.ChartSpace);

                setChartLocation(drawingsPart, chartPart, singleChart.Key);
            }
        }

        private void setChartLocation(DrawingsPart drawingsPart, ChartPart chartPart, SpredsheetLocation location)
        {
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());

            // Pozicija charta.
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId(location.ColumnIndex.ToString()),
                new ColumnOffset("581025"),
                new RowId(location.RowIndex.ToString()),
                new RowOffset("114300")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId((location.ColumnIndex + 19).ToString()),
                new ColumnOffset("276225"),
                new RowId((location.RowIndex + 15).ToString()),
                new RowOffset("0")));

            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
                    Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
                    Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            // Ime charta.
            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            twoCellAnchor.Append(new ClientData());
        }

        private double _maxWidthOfFontChar = 7d;
        private void writeColumns(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Columns());
            for (int i = 1; i <= _maxColumnIndex; i++)
            {
                var width = 20d;
                if (_maxNumberOfCharsPerColumn.ContainsKey(i))
                {
                    width = _maxNumberOfCharsPerColumn[i];
                }
                var minAtt = new OpenXmlAttribute("min", null, i.ToString());
                var maxAtt = new OpenXmlAttribute("max", null, i.ToString());
                var widthAtt = new OpenXmlAttribute("width", null, width.ToString());
                writer.WriteStartElement(new Column(), new List<OpenXmlAttribute>() {minAtt, maxAtt, widthAtt});
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public int LastRowIndex
        {
            get { return _maxRowIndex; }
        }
    }
}
