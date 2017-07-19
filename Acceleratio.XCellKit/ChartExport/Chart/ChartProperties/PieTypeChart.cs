using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Acceleratio.XCellKit
{
    public class PieTypeChart : ChartPropertiesSetup
    {
        public override void ChartAndChartSeries(string title, uint seriesNumber, PlotArea plotArea,
            out OpenXmlCompositeElement chart, out OpenXmlCompositeElement chartSeries)
        {
            chart = plotArea.AppendChild<PieChart>(new PieChart());
            chart.Append(new VaryColors() { Val = true });
            chart.Append(new FirstSliceAngle() { Val = (UInt16Value)0U });

            // Create two new line series with specified name.
            chartSeries = chart.AppendChild<PieChartSeries>(new PieChartSeries(
                new Index() { Val = new UInt32Value(seriesNumber) },
                new Order() { Val = new UInt32Value(seriesNumber) },
                new SeriesText(new NumericValue() { Text = title })));
        }

        /// <summary>
        /// Hide X axis because it is not used for PieChart.
        /// </summary>
        public override CategoryAxis SetLineCategoryAxis(PlotArea plotArea, string title = "", bool hide = false)
        {
            return base.SetLineCategoryAxis(plotArea, hide: true);
        }

        /// <summary>
        /// Hide Y axis because it is not used for PieChart.
        /// </summary>
        public override ValueAxis SetValueAxis(PlotArea plotArea, string title = "", bool hide = false)
        {
            return base.SetValueAxis(plotArea, hide: true);
        }
    }
}
