using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Acceleratio.XCellKit
{
    public class BarTypeChart : ChartPropertiesSetup
    {
        public override void ChartAndChartSeries(string title, uint seriesNumber, PlotArea plotArea,
            out OpenXmlCompositeElement chart, out OpenXmlCompositeElement chartSeries)
        {
            chart = plotArea.AppendChild<BarChart>(new BarChart());
            chart.Append(new VaryColors() { Val = false });
            chart.Append(new BarDirection() { Val = BarDirectionValues.Column });
            chart.Append(new BarGrouping() { Val = BarGroupingValues.Standard });
            chart.Append(new GapWidth() { Val = (UInt16Value)75U });
            chart.Append(new Overlap() { Val = 0 });

            // Create two new line series with specified name.
            chartSeries = chart.AppendChild<BarChartSeries>(new BarChartSeries(
                new Index() { Val = new UInt32Value(seriesNumber) },
                new Order() { Val = new UInt32Value(seriesNumber) },
                new SeriesText(new NumericValue() { Text = title })));
        }

        // Override any ChartPropertiesSetup method here.
    }
}
