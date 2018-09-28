using System.Collections.Generic;

namespace SysKit.XCellKit
{
    public class ChartSettings
    {
        /// <summary>
        /// Set chart title
        /// </summary>
        public string Title { get; set; } = null;

        /// <summary>
        /// Set X Axis Title.
        /// </summary>
        public string AxisXTitle { get; set; } = null;

        /// <summary>
        /// Set Y Axis Title
        /// </summary>
        public string AxisYTitle { get; set; } = null;

        /// <summary>
        /// Set chart Height
        /// </summary>
        public int? Height { get; set; } = null;

        /// <summary>
        /// Set chart Width
        /// </summary>
        public int? Width { get; set; } = null;

        /// <summary>
        /// Set color for each series.
        /// </summary>
        public List<string> SeriesColor { get; set; } = null;

        /// <summary>
        /// Set if Legend is visible
        /// </summary>
        public bool? Legend { get; set; } = null;

        /// <summary>
        /// Set if X Axis is visible
        /// </summary>
        public bool? AxisX { get; set; } = null;

        /// <summary>
        /// Set if Y Axis is visible
        /// </summary>
        public bool? AxisY { get; set; } = null;

        /// <summary>
        /// Set Y axis format (General, Time...)
        /// </summary>
        public virtual string AxisYFormatCategory { get; set; } = null;

        /// <summary>
        /// Set Y axis format code.
        /// </summary>
        public virtual string AxisYFormatCode { get; set; } = null;
    }
}
