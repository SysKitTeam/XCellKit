using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;

namespace Acceleratio.XCellKit.ExcelExport
{
    public abstract class ChartPropertiesSetup
    {
        /// <summary>
        /// Create chart and chart sries depending on ChartType
        /// </summary>
        public abstract void ChartAndChartSeries(
            string title,
            uint seriesNumber,
            PlotArea plotArea,
            out OpenXmlCompositeElement chart,
            out OpenXmlCompositeElement chartSeries);

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
            Marker marker = new Marker();
            marker.Append(new Symbol() { Val = MarkerStyleValues.None });

            chartSeries.Append(chartShapeProperties);
            chartSeries.Append(marker);
            chartSeries.Append(new Smooth() { Val = false });

            return chartShapeProperties;
        }

        /// <summary>
        /// Create and insert data to Axis
        /// </summary>
        public virtual void SetChartAxis(OpenXmlCompositeElement lineChart, OpenXmlCompositeElement lineChartSeries, List<ChartModel> data)
        {
            bool isArgumentDate = false;
            if (data.Count > 0)
            {
                DateTime parsedDate;
                isArgumentDate = DateTime.TryParse(data[0].Argument, out parsedDate);
            }

            uint i = 0;
            // X axis
            StringLiteral stringLiteral = new StringLiteral();
            stringLiteral.Append(new PointCount() { Val = new UInt32Value((uint)data.Count) });
            NumberLiteral numberLiteralX = new NumberLiteral();
            numberLiteralX.Append(new FormatCode("mmm dd"));
            numberLiteralX.Append(new PointCount() { Val = new UInt32Value((uint)data.Count) });

            // Y axis
            NumberLiteral numberLiteralY = lineChartSeries.AppendChild<Values>(new Values()).AppendChild<NumberLiteral>(new NumberLiteral());
            numberLiteralY.Append(new FormatCode("General"));
            numberLiteralY.Append(new PointCount() { Val = new UInt32Value((uint)data.Count) });

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
                    stringLiteral.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) })
                        .Append(new NumericValue(chartModel.Argument));
                }

                numberLiteralY.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) })
                    .Append(new NumericValue(chartModel.Value.ToString()));

                i++;
            }

            var category = lineChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData());

            if (isArgumentDate)
            {
                category.Append(numberLiteralX);
            }
            else
            {
                category.Append(stringLiteral);
            }

            lineChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            lineChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });
        }

        /// <summary>
        /// Design settings for Y axis.
        /// </summary>
        /// <param name="title">Optional parameter to set axis title</param>
        /// <param name="hide">Optiional parameter to set axis visiblity</param>
        public virtual ValueAxis SetValueAxis(PlotArea plotArea, string title = "", bool hide = false)
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

            var valueAxis = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                        DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new Delete() { Val = hide },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                majorGridlines,
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.None },
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                }, new TickLabelPosition()
                {
                    Val = new EnumValue<TickLabelPositionValues>
                        (TickLabelPositionValues.NextTo)
                }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            if (title.Length > 0)
            {
                valueAxis.Append(SetTitle(title));
            }

            return valueAxis;
        }

        /// <summary>
        /// Position the chart on the worksheet using a TwoCellAnchor object and append a GraphicFrame to the TwoCellAnchor object.
        /// </summary>
        public virtual TwoCellAnchor SetTwoCellAnchor(DrawingsPart drawingsPart, ChartPart chartPart)
        {
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());

            // Pozicija charta.
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("1"),
                new ColumnOffset("581025"),
                new RowId("1"),
                new RowOffset("114300")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("20"),
                new ColumnOffset("276225"),
                new RowId("16"),
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

            return twoCellAnchor;
        }

        /// <summary>
        /// Design settings for X axis.
        /// </summary>
        /// <param name="title">Optional parameter to set axis title</param>
        /// <param name="hide">Optiional parameter to set axis visiblity</param>
        public virtual CategoryAxis SetLineCategoryAxis(PlotArea plotArea, string title = "", bool hide = false)
        {
            var categoryAxis = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
                    { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.
                        OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new Delete() { Val = hide },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new MajorTickMark() { Val = TickMarkValues.None },
                new MinorTickMark() { Val = TickMarkValues.Outside },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = new UInt32Value(48672768U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new AutoLabeled() { Val = new BooleanValue(true) },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                new LabelOffset() { Val = new UInt16Value((ushort)100) }));
            if (title.Length > 0)
            {
                categoryAxis.Append(SetTitle(title));
            }

            return categoryAxis;
        }

        /// <summary>
        /// Design settings for legend.
        /// </summary>
        public virtual void SetLegend(Chart chart)
        {
            // Add the chart Legend.
            Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Bottom) },
                new Layout()));
            legend.Append(new Overlay() { Val = false });

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });
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
                new Overlay { Val = false },
                new ChartText(new RichText(new BodyProperties(),
                    new ListStyle(),
                    paragraph)));
        }

        /// <summary>
        /// Turns string to DateTime then back to number string that excel uses for date values.
        /// </summary>
        private string CalculateExcelDate(string dateString)
        {
            DateTime date = DateTime.Parse(dateString);
            return (date - new DateTime(1899,12,30)).TotalDays.ToString();
        }
    }
}
