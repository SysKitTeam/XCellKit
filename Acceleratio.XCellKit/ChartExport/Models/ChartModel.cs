using System;

namespace Acceleratio.XCellKit
{
    public class ChartModel : IEquatable<ChartModel>
    {
        public ChartModel(string series, string argument, double value)
        {
            this.Series = series;
            this.Argument = argument;
            this.Value = value;
        }

        public string Series { get; private set; }
        public string Argument { get; private set; }
        public double Value { get; private set; }

        public bool Equals(ChartModel other)
        {
            if (Series == other.Series && Argument == other.Argument && Value == other.Value)
                return true;

            return false;
        }

        public override int GetHashCode()
        {
            int hashFirstName = Series == null ? 0 : Series.GetHashCode();
            int hashLastName = Argument == null ? 0 : Argument.GetHashCode();
            int hashValue = Value == null ? 0 : Value.GetHashCode();

            return hashFirstName ^ hashLastName ^ hashValue;
        }
    }
}
