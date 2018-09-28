using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit.SampleApp.Demos
{
    abstract class DemoBase
    {
        public string Title { get; protected set; }
        public string Description { get; protected set; }
        public abstract void Execute();

        protected DemoBase(string title, string description)
        {
            Title = title;
            Description = description;
            OutputFile = Path.Combine(Path.GetTempPath(), $"TempXCellKitFile{DateTime.Now:yy-MM-dd_HH_mm_ss}.xlsx");
        }

        public string OutputFile { get; protected set; }
    }
}
