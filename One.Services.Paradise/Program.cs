using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Services.Paradise
{
    class Program
    {
        static void Main(string[] args)
        {
            dbDataContext db = new dbDataContext();
            var numeradores = from x in db.NUMERADORES where x.NumId == 51 select x;
            foreach (var item in numeradores)
            {
                Console.WriteLine($"{item.NumCod}");
            }
            Console.ReadLine();
        }
    }
}
