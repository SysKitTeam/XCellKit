using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace SysKit.XCellKit
{
    internal class BarTypeChart : ChartPropertiesSetup
    {
        public override OpenXmlCompositeElement CreateChart(PlotArea plotArea)
        {
            var chart = plotArea.AppendChild<BarChart>(new BarChart());
            chart.Append(new BarDirection() { Val = BarDirectionValues.Column });
            chart.Append(new BarGrouping() { Val = BarGroupingValues.Clustered });
            chart.Append(new VaryColors() { Val = false });

            return chart;
        }


        public override OpenXmlCompositeElement CreateChartSeries(string title, uint seriesNumber, OpenXmlCompositeElement chart)
        {
            // Create two new line series with specified name.
            var chartSeries = chart.AppendChild<BarChartSeries>(new BarChartSeries(
                new Index() { Val = new UInt32Value(seriesNumber) },
                new Order() { Val = new UInt32Value(seriesNumber) },
                new SeriesText(new NumericValue() { Text = title })));

            return chartSeries;
        }

        // Override any ChartPropertiesSetup method/properties here.
    }
}
