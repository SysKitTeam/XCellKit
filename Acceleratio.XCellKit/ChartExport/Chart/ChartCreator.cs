using Acceleratio.XCellKit.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;

namespace Acceleratio.XCellKit
{
    internal class ChartCreator : SpredsheetChart
    {
        public ChartCreator(SpredsheetChart export)
            : base(export)
        {

        }

        /// <summary>
        /// Creates a gantt chart.
        /// </summary>
        public Chart CreateGanttDrawing(ChartSpace chartSpace)
        {
            Chart chartContainer = chartSpace.AppendChild<Chart>(new Chart());
            chartContainer.AppendChild<AutoTitleDeleted>(new AutoTitleDeleted() { Val = false });

            // Create a new clustered column chart.
            PlotArea plotArea = chartContainer.AppendChild<PlotArea>(new PlotArea());
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

            GanttTypeChart ganttChart = new GanttTypeChart(Settings);

            ganttChart.SetChartShapeProperties(barChartSeries1, visible: false);
            ganttChart.SetChartShapeProperties(barChartSeries2, colorPoints: (uint)GanttData.Count);
            ganttChart.SetChartAxis(barChartSeries1, barChartSeries2, GanttData);

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis (X axis).
            ganttChart.SetGanttCategoryAxis(plotArea);

            // Add the Value Axis (Y axis).
            ganttChart.SetGanttValueAxis(plotArea);

            chartContainer.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            return chartContainer;
        }


        /// <summary>
        /// Creates, designes and places chart into a sheet.
        /// Calls overrideable methods for displaying the chart based on ChartType.
        /// </summary>
        public Chart CreateDrawing(ChartSpace chartSpace)
        {
            Chart chartContainer = chartSpace.AppendChild<Chart>(new Chart());
            chartContainer.AppendChild<AutoTitleDeleted>(new AutoTitleDeleted() { Val = false });

            // Create a new clustered column chart.
            PlotArea plotArea = chartContainer.AppendChild<PlotArea>(new PlotArea());

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
            chartPropertySetter.SetValues(Settings);

            // Set chart title
            chartContainer.AppendChild(chartPropertySetter.SetTitle(chartPropertySetter.Title));

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
            chartPropertySetter.SetValueAxis(plotArea);

            chartPropertySetter.SetLegend(chartContainer);

            return chartContainer;
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

