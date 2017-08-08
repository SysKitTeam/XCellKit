using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Acceleratio.XCellKit.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class GanttSpredsheetChart : SpredsheetChart
    {
        public List<GanttData> GanttData { get; protected set; }

        public GanttSpredsheetChart(List<GanttData> ganttData, ChartSettings settings)
            : base(settings)
        {
            this.GanttData = ganttData.Distinct().ToList();
        }

        internal override void CreateChart(OpenXmlWriter writer, WorksheetPart part, SpredsheetLocation location)
        {
            DrawingsPart drawingsPart = part.AddNewPart<DrawingsPart>();

            writer.WriteStartElement(new Drawing() { Id = part.GetIdOfPart(drawingsPart) });
            writer.WriteEndElement();

            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });

            Chart chartContainer = chartPart.ChartSpace.AppendChild<Chart>(new Chart());
            chartContainer.AppendChild<AutoTitleDeleted>(new AutoTitleDeleted() { Val = false });

            // Create a new clustered column chart.
            PlotArea plotArea = chartContainer.AppendChild<PlotArea>(new PlotArea());
            Layout layout1 = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart());
            barChart.Append(new BarDirection() { Val = BarDirectionValues.Bar });
            barChart.Append(new BarGrouping() { Val = BarGroupingValues.Stacked });
            barChart.Append(new GapWidth() { Val = (UInt16Value)75U });
            barChart.Append(new Overlap() { Val = 100 });

            GanttTypeChart ganttChart = new GanttTypeChart(Settings);

            var groupedData = GanttData
                .GroupBy(x => x.Name);

            List<GanttDataPairedSeries> ganttDataWithSeries = new List<GanttDataPairedSeries>();

            for (int i = 0; i < groupedData.Max(x => x.Count()); i++)
            {
                // For each series create a hidden one for spacing.
                BarChartSeries barChartSeriesHidden = barChart.AppendChild<BarChartSeries>(new BarChartSeries(
                    new Index() { Val = new UInt32Value((uint)(i * 2)) },
                    new Order() { Val = new UInt32Value((uint)(i * 2)) },
                    new SeriesText(new NumericValue() { Text = "Not Active" })));

                BarChartSeries barChartSeriesValue = barChart.AppendChild<BarChartSeries>(new BarChartSeries(
                    new Index() { Val = new UInt32Value((uint)(i * 2) + 1) },
                    new Order() { Val = new UInt32Value((uint)(i * 2) + 1) },
                    new SeriesText(new NumericValue() { Text = "Time Spent" })));

                ganttChart.SetChartShapeProperties(barChartSeriesHidden, visible: false);
                ganttChart.SetChartShapeProperties(barChartSeriesValue, colorPoints: (uint)GanttData.Count);

                var ganttData = new List<GanttData>();
                foreach (var data in groupedData.Where(x => x.Count() >= i + 1))
                {
                    ganttData.Add(data.ElementAt(i));
                }

                ganttDataWithSeries.Add(new GanttDataPairedSeries()
                {
                    BarChartSeriesHidden = barChartSeriesHidden,
                    BarChartSeriesValue = barChartSeriesValue,
                    Values = ganttData
                });
            }

            ganttChart.SetChartAxis(ganttDataWithSeries, groupedData.ToList());

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis (X axis).
            ganttChart.SetGanttCategoryAxis(plotArea);

            // Add the Value Axis (Y axis).
            ganttChart.SetGanttValueAxis(plotArea, GanttData.Min(x => x.Start), GanttData.Max(x => x.End));

            chartContainer.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            ganttChart.SetChartLocation(drawingsPart, chartPart, location);
        }

        internal class GanttDataPairedSeries
        {
            public BarChartSeries BarChartSeriesHidden { get; set; }
            public BarChartSeries BarChartSeriesValue { get; set; }
            public List<GanttData> Values  { get; set; }
        }
    }
}
