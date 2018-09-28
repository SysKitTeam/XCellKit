using System.Collections.Generic;

namespace SysKit.XCellKit
{
    internal class BaseChartProperties
    {
        /// <summary>
        /// Set chart title
        /// </summary>
        public virtual string Title { get; set; } = "";

        /// <summary>
        /// Set X Axis Title.
        /// </summary>
        public virtual string AxisXTitle { get; set; } = "";

        /// <summary>
        /// Set Y Axis Title
        /// </summary>
        public virtual string AxisYTitle { get; set; } = "";

        /// <summary>
        /// Set chart Height
        /// </summary>
        public virtual int Height { get; set; }

        /// <summary>
        /// Set chart Width
        /// </summary>
        public virtual int Width { get; set; }

        /// <summary>
        /// Set color for each series.
        /// </summary>
        public virtual List<string> SeriesColor { get; set; }

        /// <summary>
        /// Set if Legend is visible
        /// </summary>
        public virtual bool Legend { get; set; } = true;

        /// <summary>
        /// Set if X Axis is visible
        /// </summary>
        public virtual bool AxisX { get; set; } = true;

        /// <summary>
        /// Set if Y Axis is visible
        /// </summary>
        public virtual bool AxisY { get; set; } = true;

        public virtual string AxisYFormatCategory { get; set; } = "General";

        public virtual string AxisYFormatCode { get; set; } = "General";
    }
}
