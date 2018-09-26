using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acceleratio.XCellKit.Tests
{
    static class Utilities
    {
        public static double GetMemoryConsumption()
        {
            var proc = Process.GetCurrentProcess();
            
            var mem = (double)proc.PrivateMemorySize64 / 1024 / 1024;
            proc.Dispose();

            return mem;

        }
    }
}
