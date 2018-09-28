using System;

namespace SysKit.XCellKit
{
    public class GanttData : IEquatable<GanttData>
    {
        public GanttData(string name, TimeSpan start, TimeSpan end)
        {
            this.Name = name;
            this.Start = start;
            this.End = end;
        }

        public string Name { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }

        public bool Equals(GanttData other)
        {
            if (Name == other.Name && Start == other.Start && End == other.End)
                return true;

            return false;
        }

        public override int GetHashCode()
        {
            int hashName = Name == null ? 0 : Name.GetHashCode();
            int hashStart = Start == null ? 0 : Start.GetHashCode();
            int hashEnd = End == null ? 0 : End.GetHashCode();

            return hashName ^ hashStart ^ hashEnd;
        }
    }
}
