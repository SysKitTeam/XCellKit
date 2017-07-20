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
        public ChartSettings Settings { get; set; }

        protected SpredsheetChart(SpredsheetChart export)
        {
            this.ChartType = export.ChartType;
            this.ChartData = export.ChartData;
            this.GanttData = export.GanttData;
            this.Settings = export.Settings;
        }

        public SpredsheetChart(ChartTypeEnum chartType, List<ChartModel> chartData, ChartSettings settings)
        {
            this.ChartType = chartType != ChartTypeEnum.Gantt ? chartType : ChartTypeEnum.Bar;
            this.ChartData = chartData.Distinct().ToList();
            this.Settings = settings;
        }

        public SpredsheetChart(List<GanttData> ganttData, ChartSettings settings)
        {
            this.GanttData = ganttData.Distinct().ToList();
            this.ChartType = ChartTypeEnum.Gantt;
            this.Settings = settings;
        }

        internal Chart CreateChart(ChartSpace chartSpace)
        {
            ChartCreator charts = new ChartCreator(this);
            return ChartType == ChartTypeEnum.Gantt ? charts.CreateGanttDrawing(chartSpace) : charts.CreateDrawing(chartSpace);
        }
    }
}
