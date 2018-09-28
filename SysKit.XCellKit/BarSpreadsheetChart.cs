using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace SysKit.XCellKit
{
    public class BarSpreadsheetChart : SpreadsheetChart
    {
        public BarSpreadsheetChart(List<ChartModel> chartData, ChartSettings settings)
            : base(chartData, settings)
        {
            this.ChartPropertySetter = new BarTypeChart();

            this.ChartPropertyValuesFromSettings();
        }
    }
}
