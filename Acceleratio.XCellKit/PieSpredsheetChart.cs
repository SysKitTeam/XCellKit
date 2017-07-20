using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace Acceleratio.XCellKit
{
    public class PieSpredsheetChart : SpredsheetChart
    {
        public PieSpredsheetChart(List<ChartModel> chartData, ChartSettings settings)
            : base(chartData, settings)
        {
            this.ChartPropertySetter = new PieTypeChart();

            this.ChartPropertyValuesFromSettings();
        }
    }
}
