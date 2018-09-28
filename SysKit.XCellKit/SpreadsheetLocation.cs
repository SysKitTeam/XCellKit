namespace SysKit.XCellKit
{
    public class SpreadsheetLocation
    {
        private int _rowIndex;
        private int _columnIndex;
        public SpreadsheetLocation(int rowIndex, int columnIndex)
        {
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
        }

        public string ColumnName
        {
            get { return SpreadsheetHelper.ExcelColumnFromNumber(_columnIndex); }
        }

        public int RowIndex
        {
            get { return _rowIndex; }
        }

        public int ColumnIndex
        {
            get { return _columnIndex; }
        }

        public override bool Equals(object obj)
        {
            var location = obj as SpreadsheetLocation;
            if (location == null)
            {
                return false;
            }

            return _rowIndex == location._rowIndex && _columnIndex == location._columnIndex;
        }

        public override int GetHashCode()
        {
            return _rowIndex.GetHashCode() ^ _columnIndex.GetHashCode();
        }
    }
}
