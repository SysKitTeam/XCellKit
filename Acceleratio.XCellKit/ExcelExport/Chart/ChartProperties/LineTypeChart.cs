using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Acceleratio.XCellKit.ExcelExport
{
    public class LineTypeChart : ChartPropertiesSetup
    {
        public override void ChartAndChartSeries(string title, uint seriesNumber, PlotArea plotArea,
            out OpenXmlCompositeElement chart, out OpenXmlCompositeElement chartSeries)
        {
            chart = plotArea.AppendChild<LineChart>(new LineChart());
            chart.Append(new Grouping() { Val = GroupingValues.Standard });
            chart.Append(new VaryColors() { Val = false });

            // Create new line series with specified name.
            chartSeries = chart.AppendChild<LineChartSeries>(new LineChartSeries(
                new Index() { Val = new UInt32Value(seriesNumber) },
                new Order() { Val = new UInt32Value(seriesNumber) },
                new SeriesText(new NumericValue() { Text = title })));
        }

        /// <summary>
        /// LineChart has a different structure for outline.
        /// </summary>
        public override ChartShapeProperties SetChartShapeProperties(OpenXmlCompositeElement chartSeries)
        {
            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
            Outline outline = new Outline() { Width = 28575, CapType = LineCapValues.Round };
            outline.Append(new SolidFill());
            outline.Append(new Round());

            chartShapeProperties.Append(new SolidFill());
            chartShapeProperties.Append(outline);
            chartShapeProperties.Append(new EffectList());
            Marker marker = new Marker();
            marker.Append(new Symbol() { Val = MarkerStyleValues.None });

            chartSeries.Append(chartShapeProperties);
            chartSeries.Append(marker);
            chartSeries.Append(new Smooth() { Val = false });

            return chartShapeProperties;
        }
    }
}
