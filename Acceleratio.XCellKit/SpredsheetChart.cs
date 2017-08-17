using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public abstract class SpredsheetChart
    {
        public List<ChartModel> ChartData { get; protected set; }
        public ChartSettings Settings { get; set; }
        internal ChartPropertiesSetup ChartPropertySetter { get; set; }

        protected void ChartPropertyValuesFromSettings()
        {
            ChartPropertySetter.Title = Settings.Title ?? ChartPropertySetter.Title;
            ChartPropertySetter.AxisX = Settings.AxisX ?? ChartPropertySetter.AxisX;
            ChartPropertySetter.AxisXTitle = Settings.AxisXTitle ?? ChartPropertySetter.AxisXTitle;
            ChartPropertySetter.AxisY = Settings.AxisY ?? ChartPropertySetter.AxisY;
            ChartPropertySetter.AxisYTitle = Settings.AxisYTitle ?? ChartPropertySetter.AxisYTitle;
            ChartPropertySetter.Height = Settings.Height ?? ChartPropertySetter.Height;
            ChartPropertySetter.Legend = Settings.Legend ?? ChartPropertySetter.Legend;
            ChartPropertySetter.SeriesColor = Settings.SeriesColor ?? ChartPropertySetter.SeriesColor;
            ChartPropertySetter.Width = Settings.Width ?? ChartPropertySetter.Width;
        }

        protected SpredsheetChart(List<ChartModel> chartData, ChartSettings settings)
        {
            this.ChartData = chartData.Distinct().ToList();
            this.Settings = settings;
        }

        protected SpredsheetChart(ChartSettings settings)
        {
            this.Settings = settings;
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
            chartContainer.AppendChild<AutoTitleDeleted>(new AutoTitleDeleted() { Val = false });

            // Create a new clustered column chart.
            PlotArea plotArea = chartContainer.AppendChild<PlotArea>(new PlotArea());

            // Set chart title
            chartContainer.AppendChild(ChartPropertySetter.SetTitle(ChartPropertySetter.Title));

            uint chartSeriesCounter = 0;
            OpenXmlCompositeElement chart = ChartPropertySetter.CreateChart(plotArea);
            foreach (var chartDataSeriesGrouped in ChartData.GroupBy(x => x.Series))
            {
                // Set chart and series depending on type.
                OpenXmlCompositeElement chartSeries = ChartPropertySetter.CreateChartSeries(chartDataSeriesGrouped.Key, chartSeriesCounter, chart);

                // Every method from chartPropertySetter can be overriden to customize chart export.
                ChartPropertySetter.SetChartShapeProperties(chartSeries);
                ChartPropertySetter.SetChartAxis(chart, chartSeries, chartDataSeriesGrouped.ToList());

                chartSeriesCounter++;
            }

            // Add the Category Axis (X axis).
            ChartPropertySetter.SetLineCategoryAxis(plotArea);

            // Add the Value Axis (Y axis).
            ChartPropertySetter.SetValueAxis(plotArea);

            ChartPropertySetter.SetLegend(chartContainer);

            ChartPropertySetter.SetChartLocation(drawingsPart, chartPart, location);
        }
    }
}
