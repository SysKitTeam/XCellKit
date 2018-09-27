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

            public bool ExhaustedAllRows { get; private set; }

            public bool MoveNext()
            {
                if (ExhaustedAllRows)
                {
                    Current = null;
                    return false;
                }
                var args = new RequestTableRowEventArgs();
                _table.RaiseRequestTableRow(args);
                if (args.Finished)
                {
                    ExhaustedAllRows = true;
                }

                Current = args.Row;

                if (args.Row != null)
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
