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
            private readonly IEnumerator<SpreadsheetRow> _internalEnumerator;
            public SpreadSheetTableStreamingEnumerator(SpreadsheetTable table, IEnumerator<SpreadsheetRow> internalEnumerator)
            {
                _table = table;
                _internalEnumerator = internalEnumerator;
            }

            public void Dispose()
            {
                _internalEnumerator?.Dispose();
            }

            public bool ExhaustedAllRows { get; private set; }

            public bool MoveNext()
            {
                if (ExhaustedAllRows)
                {
                    Current = null;
                    return false;
                }

                if (_internalEnumerator == null)
                {
                    var args = new RequestTableRowEventArgs();
                    _table.RaiseRequestTableRow(args);
                    if (args.Finished)
                    {
                        Current = null;
                    }
                    else
                    {
                        Current = args.Row;
                    }
                    
                }
                else
                {
                    var hasData = _internalEnumerator.MoveNext();
                    Current = (SpreadsheetRow)_internalEnumerator.Current;
                    if (!hasData)
                    {
                        Current = null;
                    }
                }

                if (Current != null)
                {
                    ItemsRead++;
                }
                else
                {
                    ExhaustedAllRows = true;
                    return false;
                }
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
