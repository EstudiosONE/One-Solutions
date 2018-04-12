using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Services.Paradise.Printer
{
    internal class Helper
    {
        public class Report
        {
            public ReportPageSize PageSize { get; set; }
            public ReportPage[] Pages { get; set; }
            public class ReportPageSize
            {
                public double Width { get; set; }
                public double Height { get; set; }
            }
            public class ReportPage
            {
                public object[] Data { get; set; }
            }
            public class Text
            {
                public string Data { get; set; }
                public PdfFont Font { get; set; }
                public PdfBrush Brush{ get; set; }
                public Tuple<double, double> Position { get; set; }
            }

        }

        public static void PrintTest(string ReportData)
        {
            Report report = Newtonsoft.Json.JsonConvert.DeserializeObject<Report>(ReportData);


            //Create a new PDF document.

            PdfDocument document = new PdfDocument();

            // Set the custom page size.

            document.PageSettings.Size = new SizeF((float)Cm2Px(report.PageSize.Width), (float)Cm2Px(report.PageSize.Height));

            //Add a page to the document.

            foreach (var ItemPage in report.Pages)
            {
                PdfPage page = document.Pages.Add();
                PdfGraphics graphics = page.Graphics;

                foreach (var item in ItemPage.Data)
                {
                    var obj = item.GetType().Name;
                    switch (obj)
                    {
                        case "Text":
                            graphics.DrawString(((Report.Text)item).Data, ((Report.Text)item).Font, ((Report.Text)item).Brush, new PointF(0, 0));

                            break;
                    }
                }
            }


            //Create PDF graphics for the page.


            //Set the font.

            PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 20);

            //Draw the text.


            //Save the document.

            document.Save("Output.pdf");

            //Close the document.

            document.Close(true);
        }

        internal static double Px2Cm(double px)
        {
            return px / 28.344671201814058956916099773243;
        }
        internal static double Cm2Px(double Cm)
        {
            return Cm * 28.344671201814058956916099773243;
        }
    }
}
