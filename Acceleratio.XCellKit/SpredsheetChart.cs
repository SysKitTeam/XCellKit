using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace Acceleratio.XCellKit
{
    public class SpredsheetChart
    {
        public ChartTypeEnum ChartType { get; private set; }
        public List<ChartModel> ChartData { get; private set; }
        public List<GanttData> GanttData { get; private set; }
        public string Title { get; set; }
        public string AxisTitle { get; set; } = "";

        protected SpredsheetChart(SpredsheetChart export)
        {
            this.ChartType = export.ChartType;
            this.ChartData = export.ChartData;
            this.GanttData = export.GanttData;
            this.Title = export.Title;
            this.AxisTitle = export.AxisTitle ?? "";
        }

        public SpredsheetChart(ChartTypeEnum chartType, List<ChartModel> chartData, string title, string axisTitle)
        {
            this.ChartType = chartType != ChartTypeEnum.Gantt ? chartType : ChartTypeEnum.Bar;
            this.ChartData = chartData.Distinct().ToList();
            this.Title = title;
            this.AxisTitle = axisTitle ?? "";
        }

        public SpredsheetChart(List<GanttData> ganttData, string title)
        {
            this.GanttData = ganttData.Distinct().ToList();
            this.ChartType = ChartTypeEnum.Gantt;
            this.Title = title;
        }

        public Chart CreateChart(ChartSpace chartSpace)
        {
            ChartCreator charts = new ChartCreator(this);
            return ChartType == ChartTypeEnum.Gantt ? charts.CreateGanttDrawing(chartSpace) : charts.CreateDrawing(chartSpace);
        }
    }
}
