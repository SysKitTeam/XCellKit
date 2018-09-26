using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acceleratio.XCellKit
{
    public class RequestTableRowEventArgs : EventArgs
    {
        public SpreadsheetRow Row { get; set; }
        public bool Finished { get; set; }
    }
}
