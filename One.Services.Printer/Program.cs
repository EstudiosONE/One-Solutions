using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Services.Printer
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now);
            var document = Reports.Restaurant.Pension.Generate();
            Reports.Restaurant.Pension.Print(document);
            Console.WriteLine(DateTime.Now);
            Console.ReadLine();
        }
    }
}
