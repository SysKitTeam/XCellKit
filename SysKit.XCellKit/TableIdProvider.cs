namespace SysKit.XCellKit
{
    internal class TableIdProvider
    {
        private int _nextId = 1;
        public int GetNextId()
        {
            return _nextId++;
        }
    }
}