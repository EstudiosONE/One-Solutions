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

            document.PageSettings.Size = new SizeF((float)Helper.MM2PX(80), (float)Helper.MM2PX(200));
            document.PageSettings.Margins = new PdfMargins { All = 0 };

            //Add a page to the document.

            PdfPage page = document.Pages.Add(); 

            //Create PDF graphics for the page.

            PdfGraphics graphics = page.Graphics;

            // Logo

            PdfImage image = new PdfBitmap("Logo.png");

            float PageWidth = (float)Helper.MM2PX(60);
            float PageHeight = (float)Helper.MM2PX(30);
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
            float YPosition = ((float)Helper.MM2PX(40) - myHeight) / 2;

            graphics.DrawImage(image, XPosition, YPosition, myWidth, myHeight);

            // Titulo.
            graphics.DrawRectangle(
                new PdfPen(new PdfColor(0, 0, 0), 0.5F), 
                new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(40), (float)Helper.MM2PX(60), (float)Helper.MM2PX(12)));
            graphics.DrawRectangle(
                new PdfPen(new PdfColor(0, 0, 0), 0.5F), 
                new RectangleF((float)Helper.MM2PX(11), (float)Helper.MM2PX(41), (float)Helper.MM2PX(58), (float)Helper.MM2PX(10)));
            graphics.DrawString(
                "WIFI Free", 
                new PdfTrueTypeFont("Ubuntu-Regular.ttf", 21), 
                PdfBrushes.Black,
                new RectangleF((float)Helper.MM2PX(11), (float)Helper.MM2PX(41), (float)Helper.MM2PX(58), (float)Helper.MM2PX(10)), 
                new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle));

            // Habitacion 
            graphics.DrawRectangle(
                PdfBrushes.Black,
                new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(55), (float)Helper.MM2PX(60), (float)Helper.MM2PX(5)));
            graphics.DrawString("SANTO DOMINGO", 
                new PdfTrueTypeFont("Ubuntu-Regular.ttf", 11), 
                PdfBrushes.White,
                new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(55), (float)Helper.MM2PX(60), (float)Helper.MM2PX(5)), 
                new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle));

            // Instrucciones
            graphics.DrawString(
                "INSTRUCCIONES E INFORMACION PARA CONECTARSE A LA RED GRATUITA.",
                new PdfTrueTypeFont("Ubuntu-Regular.ttf", 11), 
                PdfBrushes.Black,
                new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(62), (float)Helper.MM2PX(60), (float)Helper.MM2PX(14)),
                new PdfStringFormat(PdfTextAlignment.Justify, PdfVerticalAlignment.Top));


            graphics.DrawString(
                "1 - Conectarse a la red WIFI “El Descubrimiento”\n"+
                "2 - Se debería abrir una página web en donde se solicitarán los siguientes datos.\n"+
                "\n"+
                $"LOGIN: 12345678\n"+
                $"CONTRASEÑA: 12345678\n"+
                "\n" +
                "3 - En caso de no abrirse dicha página web, abra en un navegador http://192.168.88.1 e ingrese los datos proporcionados anteriormente.\n" +
                "4 - Para desconectarse de la red WIFI debe primero ingresar en http://192.168.88.1/logout.html y luego desconectarse de la red WIFI “El Descubrimiento”.\n"+
                "\n" +
                "Éste usuario es brindado de forma gratuita, y cuenta con conexión para 2 dispositivos simultaneamente.\n" +
                "\n" +
                "Para contratar nuestro servicio de WIFI PLUS  sirvase pasar por RECEPCIÓN",
                new PdfTrueTypeFont("Ubuntu-Regular.ttf", 10), 
                PdfBrushes.Black,
                 new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(80), (float)Helper.MM2PX(60), (float)Helper.MM2PX(114)), 
                new PdfStringFormat(PdfTextAlignment.Justify, PdfVerticalAlignment.Top));

            return document;
        }

        internal static void Print(PdfDocument document)
        {
            Helper.Print(document, @"MP-4000 TH");
        }
    }
}
