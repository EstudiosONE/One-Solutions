using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Services.Printer.Reports.Restaurant
{
    internal class Pension
    {
        internal static PdfDocument Generate()
        {
            PdfDocument document = new PdfDocument();

            document.PageSettings.Size = PdfPageSize.A4;
            document.PageSettings.Size = new SizeF((float)Helper.MM2PX(80), (float)Helper.MM2PX(200));
            document.PageSettings.Margins = new PdfMargins { All = 0 };

            //Add a page to the document.

            PdfPage page = document.Pages.Add();

            //Create PDF graphics for the page.

            PdfGraphics graphics = page.Graphics;

            // Logo

            PdfImage image = new PdfBitmap("logo.jpg");

            float PageWidth = (float)Helper.MM2PX(60);
            float PageHeight = (float)Helper.MM2PX(20);
            float myWidth = image.Width;
            float myHeight = image.Height;

            float shrinkFactor;

            if (myWidth > PageWidth)
            {
                shrinkFactor = myWidth / PageWidth;
                myWidth = PageWidth;
                myHeight = myHeight / shrinkFactor;
            }

            if (myHeight > PageHeight)
            {
                shrinkFactor = myHeight / PageHeight;
                myHeight = PageHeight;
                myWidth = myWidth / shrinkFactor;
            }

            float XPosition = ((float)Helper.MM2PX(80) - myWidth) / 2;
            float YPosition = ((float)Helper.MM2PX(30) - myHeight) / 2;

            graphics.DrawImage(image, XPosition, YPosition, myWidth, myHeight);


            // Titulo.
            var TitleRectangle = new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(30), (float)Helper.MM2PX(60), (float)Helper.MM2PX(20));
            graphics.DrawRectangle(new PdfPen(new PdfColor(0, 0, 0), 1), TitleRectangle);
            graphics.DrawString("WIFI Free", new PdfTrueTypeFont("Lato-Bold.ttf", 24), PdfBrushes.Black, TitleRectangle, new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle));

            // Subtitulo
            var SubTitleRectangle = new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(55), (float)Helper.MM2PX(60), (float)Helper.MM2PX(15));
            graphics.DrawString("Instrucciones para conectarse correctamente a WIFI Free:", new PdfTrueTypeFont("Lato-Bold.ttf", 12, PdfFontStyle.Underline), PdfBrushes.Black, SubTitleRectangle, new PdfStringFormat(PdfTextAlignment.Left, PdfVerticalAlignment.Middle));

            // Habitacion 
            var HabitacionRectangle = new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(75), (float)Helper.MM2PX(60), (float)Helper.MM2PX(20));
            graphics.DrawString("Prueba", new PdfTrueTypeFont("Lato-Bold.ttf", 24), PdfBrushes.Black, HabitacionRectangle, new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle));

            // Instrucciones

            var InstructionsRectangle = new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(95), (float)Helper.MM2PX(60), (float)Helper.MM2PX(80));
            graphics.DrawString(
                "1.- Conectarse a la red \"EL DESCUBRIMIENTO\".\n" +
                "2.- Si se abre automáticamente una pagina para ingresar usuario y contraseña ingrese los datos que a continuación se proporcionan.\n" +
                "\n" +
                $"Usuario: 1234\n" +
                $"Contraseña: 1234\n" +
                "\n" +
                "3.- Si no abre ninguna página automáticamente, puede ir a \"http://hotspot.info\" e ingresar los datos proporcionados anteriormente.\n" +
                "4.- Para desconectar el equipo de la red ingresar a \"http://hotspot.info\" y presionar el botón \"DESCONECTAR\"",
                new PdfTrueTypeFont("Lato-Regular.ttf", 10), PdfBrushes.Black, InstructionsRectangle, new PdfStringFormat(PdfTextAlignment.Left, PdfVerticalAlignment.Top));

            return document;
        }

        internal static void Print(PdfDocument document)
        {
            Helper.Print(document, @"\\minimercado01\eFACT");
        }
    }
}
