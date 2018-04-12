using Syncfusion.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static One.Services.Paradise.Printer.Helper;

namespace One.Services.Paradise
{
    class Program
    {
        static void Main(string[] args)
        {
            Report report = new Report()
            {
                PageSize = new Report.ReportPageSize() { Width = 8, Height = 15 },
                Pages = new Report.ReportPage[]
                {
                    new Report.ReportPage()
                    {
                        Data = new object[]
                        {
                            new Report.Text()
                            {
                                Data = "Hola",
                                Font =new PdfStandardFont(PdfFontFamily.Helvetica, 20),
                                Brush = PdfBrushes.DarkRed
                            }
                        }
                    }
                }
            };
            PrintTest(Newtonsoft.Json.JsonConvert.SerializeObject(report));
            Console.ReadLine();
        }
    }
}
