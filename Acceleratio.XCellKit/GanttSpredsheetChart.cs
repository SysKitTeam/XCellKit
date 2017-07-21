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
            this.GanttData = ganttData;
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

            ganttChart.SetChartLocation(drawingsPart, chartPart, location);
        }
    }
}
