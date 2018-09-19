using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public abstract class SpredsheetChart
    {
        public List<ChartModel> ChartData { get; protected set; }
        public ChartSettings UserSettings { get; set; }
        internal ChartPropertiesSetup ChartPropertySetter { get; set; }

        protected void ChartPropertyValuesFromSettings()
        {
            foreach (PropertyInfo pi in typeof(ChartSettings).GetProperties())
            {
                if (pi.GetValue(UserSettings, null) != null)
                {
                    typeof(BaseChartProperties)
                        .GetProperty(pi.Name)
                        .SetValue(ChartPropertySetter.ChartProperties, pi.GetValue(UserSettings, null));
                }
            }
        }

        protected SpredsheetChart(List<ChartModel> chartData, ChartSettings settings)
        {
            this.ChartData = chartData.Distinct().ToList();
            this.UserSettings = settings;
        }

        protected SpredsheetChart(ChartSettings settings)
        {
            this.UserSettings = settings;
        }

        internal virtual void CreateChart(OpenXmlWriter writer, WorksheetPart part, SpredsheetLocation location)
        {
            DrawingsPart drawingsPart = part.AddNewPart<DrawingsPart>();

            writer.WriteStartElement(new Drawing() { Id = part.GetIdOfPart(drawingsPart) });
            writer.WriteEndElement();

            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });


            Chart chartContainer = chartPart.ChartSpace.AppendChild<Chart>(new Chart());
            // Set chart title
            chartContainer.AppendChild(ChartPropertySetter.SetTitle(ChartPropertySetter.ChartProperties.Title));
            chartContainer.AppendChild<AutoTitleDeleted>(new AutoTitleDeleted() { Val = false });

            // Create a new clustered column chart.
            PlotArea plotArea = chartContainer.AppendChild<PlotArea>(new PlotArea());

            uint chartSeriesCounter = 0;
            OpenXmlCompositeElement chart = ChartPropertySetter.CreateChart(plotArea);
            foreach (var chartDataSeriesGrouped in ChartData.GroupBy(x => x.Series))
            {
                // Set chart and series depending on type.
                OpenXmlCompositeElement chartSeries = ChartPropertySetter.CreateChartSeries(chartDataSeriesGrouped.Key, chartSeriesCounter, chart);

                // Every method from chartPropertySetter can be overriden to customize chart export.
                ChartPropertySetter.SetChartShapeProperties(chartSeries);
                ChartPropertySetter.SetChartAxis(chartSeries, chartDataSeriesGrouped.ToList());

                chartSeriesCounter++;
            }

            chart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            chart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis (X axis).
            ChartPropertySetter.SetLineCategoryAxis(plotArea);

            // Add the Value Axis (Y axis).
            ChartPropertySetter.SetValueAxis(plotArea);

            ChartPropertySetter.SetLegend(chartContainer);

            ChartPropertySetter.SetChartLocation(drawingsPart, chartPart, location);
        }
    }
}
