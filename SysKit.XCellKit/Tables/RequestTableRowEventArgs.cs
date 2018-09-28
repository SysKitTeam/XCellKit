using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acceleratio.XCellKit
{
    public class RequestTableRowEventArgs : EventArgs
    {
        private bool _finished;
        public SpreadsheetRow Row { get; set; }

        public bool Finished
        {
            get
            {
                if (Row == null)
                {
                    return true;
                }
                return _finished;
            }
            set => _finished = value;
        }
    }
}
