using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acceleratio.XCellKit
{

    public partial class SpreadsheetTable
    {
        public class SpreadSheetTableStreamingEnumerator : IEnumerator<SpreadsheetRow>
        {
            private readonly SpreadsheetTable _table;

            public SpreadSheetTableStreamingEnumerator(SpreadsheetTable table)
            {
                _table = table;
            }

            public void Dispose()
            {
            }

            private bool _exhaustedAllRows;
            public bool MoveNext()
            {
                if (_exhaustedAllRows)
                {
                    return false;
                }
                var args = new RequestTableRowEventArgs();
                _table.RaiseRequestTableRow(args);
                if (args.Finished)
                {
                    _exhaustedAllRows = true;
                    Current = null;
                    return false;
                }

                Current = args.Row;
                ItemsRead++;
                return true;
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }

            public SpreadsheetRow Current { get; private set; }

            object IEnumerator.Current
            {
                get { return Current; }
            }

            public int ItemsRead { get; private set; }
        }
    }
}
