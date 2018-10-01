using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using SysKit.XCellKit.SampleApp.Demos;

namespace SysKit.XCellKit.SampleApp
{
    class MenuSelection
    {
        public bool Quit { get; set; }
        public int? DemoIndex { get; set; }
        public bool IsValid { get; set; }
    }
    class Program
    {
        static List<DemoBase> _allDemos = new List<DemoBase>();
        static void Main(string[] args)
        {
            readAllDemos();

            while (true)
            {
                renderMenu();
                var selection = chooseMenuItem();
                if (selection.Quit)
                {
                    break;
                }

                if (selection.IsValid)
                {
                    executeDemo(selection.DemoIndex.Value);
                }
            }
        }

        private static void executeDemo(int idx)
        {
            Console.Clear();
            Console.WriteLine("Generating...");
            var demo = _allDemos[idx];
            demo.Execute();
            Console.Clear();
            var p = Process.Start(demo.OutputFile);
            Console.WriteLine("Waiting for excel exit...");
            p.WaitForExit();
           
            
        }


        private static MenuSelection chooseMenuItem()
        {
            Console.Write("Input selection and press enter to confirm: ");
            var selection = Console.ReadLine();
            if (string.Equals(selection, "q", StringComparison.OrdinalIgnoreCase))
            {
                return new MenuSelection() { Quit = true };
            }

            if (!Int32.TryParse(selection, out int menuItem)
            || menuItem  < 1
            || menuItem > _allDemos.Count)
            {
                return new MenuSelection() { IsValid = false };
            }

            return new MenuSelection() { IsValid = true, DemoIndex = menuItem - 1 };

        }

        private static void readAllDemos()
        {
            _allDemos = Assembly.GetExecutingAssembly()
                .GetTypes()
                .Where(x => typeof(DemoBase).IsAssignableFrom(x) && !x.IsAbstract)
                .Select(x => (DemoBase)Activator.CreateInstance(x))
                .ToList();
        }

        private static void renderMenu()
        {
            Console.Clear();
            for (var i = 0; i < _allDemos.Count; i++)
            {
                var demo = _allDemos[i];
                Console.WriteLine($"{i + 1} - {demo.Title}");
            }
            Console.WriteLine("q - Quit");
        }
    }
}
