using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;
using Acceleratio.XCellKit.ExcelExport.Helpers;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;

namespace Acceleratio.XCellKit.ExcelExport
{
    public class Charts : ReportExport
    {
        public Charts(ReportExport export)
            : base(export)
        {

        }
        /// <summary>
        /// Creates, designes and places chart into a sheet.
        /// Calls overrideable methods for displaying the chart based on ChartType.
        /// </summary>
        public void InsertChartInSpreadsheet(Stream stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                #region Chart Container - Same for all chart types
                WorksheetPart worksheetPart = createNewSheet(document, "Graph");

                // Add a new drawing to the worksheet.
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new Drawing()
                { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();

                // Add a new chart and set the chart language to English-US.
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace(new EditingLanguage() { Val = new StringValue("en-US") });
                Chart chartContainer = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart());
                chartContainer.AppendChild<AutoTitleDeleted>(new AutoTitleDeleted() { Val = false });

                // Create a new clustered column chart.
                PlotArea plotArea = chartContainer.AppendChild<PlotArea>(new PlotArea());
                #endregion

                ChartPropertiesSetup chartPropertySetter;
                switch (ChartType)
                {
                    case ChartTypeEnum.Bar:
                        chartPropertySetter = new BarTypeChart();
                        break;
                    case ChartTypeEnum.Line:
                        chartPropertySetter = new LineTypeChart();
                        break;
                    case ChartTypeEnum.Pie:
                        chartPropertySetter = new PieTypeChart();
                        break;
                    default:
                        chartPropertySetter = new BarTypeChart();
                        break;
                }

                // Set chart title
                chartContainer.AppendChild<Title>(chartPropertySetter.SetTitle(Title));

                uint chartSeriesCounter = 0;
                foreach (var chartDataSeriesGrouped in ChartData.GroupBy(x => x.Series))
                {
                    // CHART BODY - Depends on chart type.
                    OpenXmlCompositeElement chart;
                    OpenXmlCompositeElement chartSeries;

                    // Set chart and series depending on type.
                    chartPropertySetter.ChartAndChartSeries(chartDataSeriesGrouped.Key, chartSeriesCounter, plotArea, out chart, out chartSeries);

                    // Every method from chartPropertySetter can be overriden to customize chart export.
                    chartPropertySetter.SetChartShapeProperties(chartSeries);
                    chartPropertySetter.SetChartAxis(chart, chartSeries, chartDataSeriesGrouped.ToList());

                    chartSeriesCounter++;
                }

                // Add the Category Axis (X axis).
                chartPropertySetter.SetLineCategoryAxis(plotArea);

                // Add the Value Axis (Y axis).
                chartPropertySetter.SetValueAxis(plotArea, title: AxisTitle);

                chartPropertySetter.SetLegend(chartContainer);

                // Save the chart part.
                chartPart.ChartSpace.Save();

                // Position the chart on the worksheet using a TwoCellAnchor object and append a GraphicFrame to the TwoCellAnchor object..
                chartPropertySetter.SetTwoCellAnchor(drawingsPart, chartPart);

                // END - Save the WorksheetDrawing object.
                drawingsPart.WorksheetDrawing.Save();
            }
        }

        /// <summary>
        /// Creates a gantt chart.
        /// </summary>
        public void InsertGanttChartInSpreadsheet(Stream stream, string worksheetName)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                // Create new sheet with sent name. 
                WorksheetPart worksheetPart = createNewSheet(document, worksheetName);

                // Add a new drawing to the worksheet.
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
                { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();

                // Add a new chart and set the chart language to English-US.
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
                DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart());

                // Create a new clustered column chart.
                PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
                Layout layout1 = plotArea.AppendChild<Layout>(new Layout());
                BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart());
                barChart.Append(new BarDirection() { Val = BarDirectionValues.Bar });
                barChart.Append(new BarGrouping() { Val = BarGroupingValues.Stacked });
                barChart.Append(new GapWidth() { Val = (UInt16Value)75U });
                barChart.Append(new Overlap() { Val = 100 });

                // Create two new line series with specified name.
                BarChartSeries barChartSeries1 = barChart.AppendChild<BarChartSeries>(new BarChartSeries(
                    new Index() { Val = new UInt32Value(0u) },
                    new Order() { Val = new UInt32Value(0u) },
                    new SeriesText(new NumericValue() { Text = "Start Date" })));

                BarChartSeries barChartSeries2 = barChart.AppendChild<BarChartSeries>(new BarChartSeries(
                    new Index() { Val = new UInt32Value(1u) },
                    new Order() { Val = new UInt32Value(1u) },
                    new SeriesText(new NumericValue() { Text = "Time Spent" })));

                GanttTypeChart.SetChartShapeProperties(barChartSeries1, visible: false);
                GanttTypeChart.SetChartShapeProperties(barChartSeries2, colorPoints: (uint)GanttData.Count);
                GanttTypeChart.SetChartAxis(barChartSeries1, barChartSeries2, GanttData);

                barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
                barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

                // Add the Category Axis (X axis).
                GanttTypeChart.SetGanttCategoryAxis(plotArea);

                // Add the Value Axis (Y axis).
                GanttTypeChart.SetGanttValueAxis(plotArea);

                chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

                // Save the chart part.
                chartPart.ChartSpace.Save();

                // Position the chart on the worksheet using a TwoCellAnchor object and append a GraphicFrame to the TwoCellAnchor object..
                GanttTypeChart.TwoCellAnchor(drawingsPart, chartPart);

                // Save the WorksheetDrawing object.
                drawingsPart.WorksheetDrawing.Save();
            }
        }

        /// <summary>
        /// Helper method to create a new sheet.
        /// </summary>
        private WorksheetPart createNewSheet(SpreadsheetDocument document, string worksheetName)
        {
            WorksheetPart worksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = document.WorkbookPart.GetIdOfPart(worksheetPart);

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = 2, Name = worksheetName };
            sheets.Append(sheet);

            return worksheetPart;
        }
    }
}

