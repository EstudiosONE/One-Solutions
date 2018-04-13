using Syncfusion.Pdf;
using Syncfusion.Windows.Forms.PdfViewer;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace One.Services.Printer.Reports
{
    internal class Helper
    {
        // Variables estaticas
        static double dpi;

        // Inicializacion estática del Helper
        static Helper()
        {
            PdfDocument document = new PdfDocument();

            document.PageSettings.Size = PdfPageSize.A4;

            dpi = document.PageSettings.Size.Width / 210;
        }

        // Propiedades de acceso estático
        internal static double DPI { get => dpi; }

        // Métodos
        internal static double MM2PX(double MM)
        {
            return MM * DPI;
        }

        internal static double PX2MM(double PX)
        {
            return PX / DPI;
        }

        internal static double MM2IN(double MM)
        {
            return MM * 0.0393701;
        }

        internal static double IN2MM(double IN)
        {
            return IN / 0.0393701;
        }

        internal static double PX2IN(double PX)
        {
            var MM = PX2MM(PX);
            var IN = MM2IN(MM);
            return IN;
        }

        internal static double IN2PX(double IN)
        {
            return MM2PX(IN2MM(IN));
        }

        internal static void Print(PdfDocument document, string printerName)
        {
            // Declaracion de variables
            PdfViewerControl viewer;
            PrintDocument printDocument;
            PrintDialog dialog;

            viewer = new PdfViewerControl();
            dialog = new PrintDialog();

            document.Save("Output.pdf");
            viewer.Load("Output.pdf");

            printDocument = viewer.PrintDocument;
            printDocument.DefaultPageSettings.Landscape = document.PageSettings.Orientation == PdfPageOrientation.Landscape ? true : false;

            printDocument.DefaultPageSettings.Margins = new Margins(
                Convert.ToInt32(document.PageSettings.Margins.Left),
                Convert.ToInt32(document.PageSettings.Margins.Right),
                Convert.ToInt32(document.PageSettings.Margins.Top),
                Convert.ToInt32(document.PageSettings.Margins.Bottom));
            var paperW = PX2IN(document.PageSettings.Size.Width);
            var IpaperW = Convert.ToInt16(paperW * 100);

            printDocument.DefaultPageSettings.PaperSize = new PaperSize("DEF", Convert.ToInt16(PX2IN(document.PageSettings.Size.Width) * 100), Convert.ToInt16(PX2IN(document.PageSettings.Size.Height) * 100));

            dialog.Document = printDocument;

            dialog.Document.PrinterSettings.PrinterName = printerName;

            dialog.Document.Print();
        }
    }
}
