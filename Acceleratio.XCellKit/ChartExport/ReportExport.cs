using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Acceleratio.XCellKit
{
    public class ReportExport
    {
        public ChartTypeEnum ChartType { get; private set; }
        public DataTable ReportData { get; private set; }
        public List<ChartModel> ChartData { get; private set; }
        public List<GanttData> GanttData { get; private set; }
        public string Title { get; set; }
        public string AxisTitle { get; set; } = "";

        protected ReportExport(ReportExport export)
        {
            this.ChartType = export.ChartType;
            this.ChartData = export.ChartData;
            this.GanttData = export.GanttData;
            this.ReportData = export.ReportData;
            this.Title = export.Title;
            this.AxisTitle = export.AxisTitle ?? "";
        }

        public ReportExport(DataTable reportData, ChartTypeEnum chartType, List<ChartModel> chartData, string title, string axisTitle)
        {
            this.ReportData = reportData;
            this.ChartType = chartType != ChartTypeEnum.Gantt ? chartType : ChartTypeEnum.Bar;
            this.ChartData = chartData.Distinct().ToList();
            this.Title = title;
            this.AxisTitle = axisTitle ?? "";
        }

        public ReportExport(DataTable reportData, List<GanttData> ganttData, string title)
        {
            this.ReportData = reportData;
            this.GanttData = ganttData.Distinct().ToList();
            this.ChartType = ChartTypeEnum.Gantt;
            this.Title = title;
        }

        //public void CreateExcel()
        //{
        //    try
        //    {
        //        using (var stream = ExportData.OpenEntryStreamForWriting())
        //        {
        //            if (ChartType == ChartTypeEnum.Gantt)
        //            {
        //                ExportData.InsertData(stream, ReportData);

        //                ChartCreator charts = new ChartCreator(this);
        //                charts.InsertGanttChartInSpreadsheet(stream, "Graph");
        //            }
        //            else
        //            {
        //                ExportData.InsertData(stream, ReportData);
        //                ChartCreator charts = new ChartCreator(this);
        //                charts.InsertChartInSpreadsheet(stream);
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        var test = e;
        //    }
        //}
    }
}
