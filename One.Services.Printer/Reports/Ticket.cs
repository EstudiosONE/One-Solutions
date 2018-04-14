using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Services.Printer.Reports
{
    internal partial class Ticket
    {
        /// <summary>
        /// Genera un nuevo documento PDF adaptado al formato Ticket
        /// </summary>
        /// <returns>Documento PDF en formato 80 x 300</returns>
        internal static PdfDocument NewDocument()
        {
            PdfDocument document = new PdfDocument();

            document.PageSettings.Size = new SizeF((float)Helper.MM2PX(80), (float)Helper.MM2PX(300));
            document.PageSettings.Margins = new PdfMargins { All = 0 };

            return document;
        }
        /// <summary>
        /// Agrega una nueva página al documento especificado
        /// </summary>
        /// <param name="document">Documento PDF</param>
        /// <returns>Documento PDF modificado</returns>
        internal static PdfDocument AddPage(PdfDocument document)
        {
            document.Pages.Add();

            return document;
        }
        /// <summary>
        /// Genera el Header en el documento especificado
        /// </summary>
        /// <param name="document">Documento PDF</param>
        /// <param name="title">Titulo del reporte</param>
        /// <returns>Documento PDF modificado</returns>
        internal static PdfDocument GenerateHeader(PdfDocument document, string title)
        {
            foreach (PdfPage page in document.Pages)
            {
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
                    title,
                    new PdfTrueTypeFont("Ubuntu-Regular.ttf", 21),
                    PdfBrushes.Black,
                    new RectangleF((float)Helper.MM2PX(11), (float)Helper.MM2PX(41), (float)Helper.MM2PX(58), (float)Helper.MM2PX(10)),
                    new PdfStringFormat(PdfTextAlignment.Center, PdfVerticalAlignment.Middle));
            }

            return document;
        }

    }
}
