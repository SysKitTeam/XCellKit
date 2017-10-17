using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Boolean = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle.Boolean;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;

namespace Acceleratio.XCellKit
{
    internal abstract class ChartPropertiesSetup
    {
#region Properties

        public virtual BaseChartProperties ChartProperties { get; set; } = new BaseChartProperties();

        private bool isArgumentDate { get; set; } = false;

        private int dataCount { get; set; } = 0;

        private double yAxisValue { get; set; } = 0;
#endregion

        /// <summary>
        /// Create chart and chart sries depending on ChartType
        /// </summary>
        public abstract OpenXmlCompositeElement CreateChart(PlotArea plotArea);

        public abstract OpenXmlCompositeElement CreateChartSeries(
            string title,
            uint seriesNumber,
            OpenXmlCompositeElement chart);

        /// <summary>
        /// Set display, width, color and fill of borders and data (line, bar etc.) in chart.
        /// </summary>
        public virtual ChartShapeProperties SetChartShapeProperties(OpenXmlCompositeElement chartSeries)
            
        {
            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
            Outline outline = new Outline() { Width = 28575, CapType = LineCapValues.Round };
            outline.Append(new NoFill());
            outline.Append(new Round());

            chartShapeProperties.Append(outline);
            chartShapeProperties.Append(new EffectList());

            chartSeries.Append(chartShapeProperties);

            return chartShapeProperties;
        }

        /// <summary>
        /// Create and insert data to Axis
        /// </summary>
        public virtual void SetChartAxis(OpenXmlCompositeElement chartSeries, List<ChartModel> data)
        {
            dataCount = data.Count;
            if (dataCount > 0 && !isArgumentDate)
            {
                DateTime parsedDate;
                isArgumentDate = DateTime.TryParse(data[0].Argument, out parsedDate);
            }

            uint i = 0;
            // X axis
            StringLiteral stringLiteral = new StringLiteral();
            stringLiteral.Append(new PointCount() { Val = new UInt32Value((uint)dataCount) });
            NumberLiteral numberLiteralX = new NumberLiteral();
            numberLiteralX.Append(new FormatCode("mmm dd"));
            numberLiteralX.Append(new PointCount() { Val = new UInt32Value((uint)dataCount) });

            // Y axis
            NumberLiteral numberLiteralY = new NumberLiteral();
            numberLiteralY.Append(new FormatCode(ChartProperties.AxisYFormatCode));
            numberLiteralY.Append(new PointCount() { Val = new UInt32Value((uint)dataCount) });

            yAxisValue =  data.Max(x => x.Value) > yAxisValue ? data.Max(x => x.Value) : yAxisValue;

            // Set values to X and Y axis.
            foreach (var chartModel in data)
            {
                if (isArgumentDate)
                {
                    numberLiteralX.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) })
                        .Append(new NumericValue(CalculateExcelDate(chartModel.Argument)));
                }
                else
                {
                    stringLiteral.Append(new StringPoint() { Index = new UInt32Value(i), NumericValue = new NumericValue() { Text = chartModel.Argument } });
                }

                numberLiteralY.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) })
                    .Append(new NumericValue(ChartProperties.AxisYFormatCategory == "Time" ? ((double)chartModel.Value / 86400).ToString(System.Globalization.CultureInfo.InvariantCulture) : chartModel.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)));

                i++;
            }

            if (isArgumentDate)
            {
                chartSeries.Append(new CategoryAxisData() { NumberLiteral = numberLiteralX });
            }
            else
            {
                chartSeries.Append(new CategoryAxisData() { StringLiteral = stringLiteral });
            }

            chartSeries.Append(new Values() { NumberLiteral = numberLiteralY });
        }

        /// <summary>
        /// Design settings for Y axis.
        /// </summary>
        public virtual ValueAxis SetValueAxis(PlotArea plotArea)
        {
            // Postavljanje Gridline-a.
            MajorGridlines majorGridlines = new MajorGridlines();
            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
            Outline outline = new Outline();
            SolidFill solidFill = new SolidFill();
            SchemeColor schemeColor = new SchemeColor() { Val = SchemeColorValues.Accent1 };
            Alpha alpha = new Alpha() { Val = 10000 };
            schemeColor.Append(alpha);
            solidFill.Append(schemeColor);
            outline.Append(solidFill);
            chartShapeProperties.Append(outline);
            majorGridlines.Append(chartShapeProperties);

            var valueAxis = plotArea.AppendChild<ValueAxis>(new ValueAxis(
                new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                        DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new Delete() { Val = !ChartProperties.AxisY },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                majorGridlines,
                SetTitle(ChartProperties.AxisYTitle),
                new NumberingFormat() {
                    FormatCode = ChartProperties.AxisYFormatCode,
                    SourceLinked = new BooleanValue(true)
                },
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.None },
                new TickLabelPosition()
                {
                    Val = new EnumValue<TickLabelPositionValues>
                        (TickLabelPositionValues.NextTo)
                }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            if (ChartProperties.AxisYFormatCategory == "Time")
            {
                valueAxis.Append(new MajorUnit() { Val = getMajorUnitFromSeconds((int)yAxisValue) });
            }

            return valueAxis;
        }

        /// <summary>
        /// Design settings for X axis.
        /// </summary>
        /// <param name="title">Optional parameter to set axis title</param>
        /// <param name="hide">Optiional parameter to set axis visiblity</param>
        public virtual OpenXmlElement SetLineCategoryAxis(PlotArea plotArea)
        {
            List<OpenXmlElement> axisChildElements = new List<OpenXmlElement> () {
                new AxisId() {Val = new UInt32Value(48650112u)},
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.
                        OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts
                        .OrientationValues.MinMax)
                }),
                new Delete() {Val = !ChartProperties.AxisX},
                new AxisPosition() {Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom)},
                new NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new MajorTickMark() {Val = TickMarkValues.None},
                new MinorTickMark() {Val = TickMarkValues.Outside},
                new TickLabelPosition() {Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)},
                new CrossingAxis() {Val = new UInt32Value(48672768U)},
                new Crosses() {Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero)},
                new AutoLabeled() {Val = new BooleanValue(true)},
                new LabelAlignment() {Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center)},
                new LabelOffset() {Val = new UInt16Value((ushort) 100)}
            };

            if (this.isArgumentDate)
            {
                axisChildElements.Add(new MajorUnit() {Val = dataCount > 20 ? (int) (dataCount / 10) : 1});
            }

            var categoryAxis = isArgumentDate ? (OpenXmlElement) plotArea.AppendChild(new DateAxis(axisChildElements)) : plotArea.AppendChild(new CategoryAxis(axisChildElements));

            if (ChartProperties.AxisXTitle.Length > 0)
            {
                categoryAxis.Append(SetTitle(ChartProperties.AxisXTitle));
            }

            return categoryAxis;
        }

        /// <summary>
        /// Design settings for legend.
        /// </summary>
        public virtual void SetLegend(Chart chart)
        {
            if (ChartProperties.Legend)
            {
                // Add the chart Legend.
                Legend legend = chart.AppendChild<Legend>(
                    new Legend(
                        new LegendPosition() {Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Bottom)},
                        new Layout()));
                legend.Append(new Overlay() {Val = false});

                chart.Append(new PlotVisibleOnly() {Val = new BooleanValue(true)});
            }
        }

        /// <summary>
        /// Set title to parent. Used by Chart and Axis.
        /// </summary>
        public virtual Title SetTitle(string titleText)
        {
            Paragraph paragraph = new Paragraph(
                new ParagraphProperties(new DefaultRunProperties()),
                new Run(new RunProperties(),
                    new Text { Text = titleText }));

            return new Title(
                new ChartText(new RichText(new BodyProperties(),
                    new ListStyle(),
                    paragraph)),
                new Overlay {Val = false});
        }

        public void SetChartLocation(DrawingsPart drawingsPart, ChartPart chartPart, SpredsheetLocation location)
        {
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());

            // Pozicija charta.
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId(location.ColumnIndex.ToString()),
                new ColumnOffset("0"),
                new RowId(location.RowIndex.ToString()),
                new RowOffset("114300")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId((location.ColumnIndex + 19).ToString()),
                new ColumnOffset("0"),
                new RowId((location.RowIndex + 15).ToString()),
                new RowOffset("0")));

            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
                    Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
                    Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            // Ime charta.
            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            twoCellAnchor.Append(new ClientData());
        }


        /// <summary>
        /// Turns string to DateTime then back to number string that excel uses for date values.
        /// </summary>
        private string CalculateExcelDate(string dateString)
        {
            DateTime date = DateTime.Parse(dateString);
            return (date - new DateTime(1899,12,30)).TotalDays.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private double getMajorUnitFromSeconds(int maxSeconds)
        {
            var hPoTick = (double)maxSeconds / 5 / 3600;

            var value = hPoTick;
            var tick = 0;
            while (value > 10)
            {
                tick++;
                value = value / 10;
            }

            var end = value <= 1 ? 1 : value <= 2 ? 2 : value <= 5 ? 5 : 10;
            return (double)(end * Math.Pow(10, tick)) / 24;
        }
    }
}
