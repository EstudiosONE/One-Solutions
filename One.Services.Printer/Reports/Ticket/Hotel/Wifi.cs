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
        internal partial class Hotel
        {
            internal class Wifi
            {
                PdfDocument document;

                /// <summary>
                /// Genera un nuevo reporte de WIFI
                /// </summary>
                /// <param name="wifi">Datos del reporte</param>
                public Wifi(WifiType wifi)
                {
                    // Generar el nuevo documento
                    document = NewDocument();
                    var page = document.Pages.Add();
                    var graphics = page.Graphics;

                    // Habitacion 
                    graphics.DrawRectangle(
                        PdfBrushes.Black,
                        new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(55), (float)Helper.MM2PX(60), (float)Helper.MM2PX(5)));
                    graphics.DrawString(wifi.Room,
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
                        "1 - Conectarse a la red WIFI “El Descubrimiento”\n" +
                        "2 - Se debería abrir una página web en donde se solicitarán los siguientes datos.\n" +
                        "\n" +
                        $"LOGIN: {wifi.Login}\n" +
                        $"CONTRASEÑA: {wifi.Pass}\n" +
                        "\n" +
                        "3 - En caso de no abrirse dicha página web, abra en un navegador http://192.168.88.1 e ingrese los datos proporcionados anteriormente.\n" +
                        "4 - Para desconectarse de la red WIFI debe primero ingresar en http://192.168.88.1/logout.html y luego desconectarse de la red WIFI “El Descubrimiento”.\n" +
                        "\n" +
                        $"Éste usuario es brindado de forma gratuita, y cuenta con conexión para {wifi.FreeUsers} dispositivos simultaneamente.\n" +
                        "\n" +
                        "Para contratar nuestro servicio de WIFI PLUS  sirvase pasar por RECEPCIÓN",
                        new PdfTrueTypeFont("Ubuntu-Regular.ttf", 10),
                        PdfBrushes.Black,
                        new RectangleF((float)Helper.MM2PX(10), (float)Helper.MM2PX(80), (float)Helper.MM2PX(60), (float)Helper.MM2PX(114)),
                        new PdfStringFormat(PdfTextAlignment.Justify, PdfVerticalAlignment.Top));

                    // Header
                    document = GenerateHeader(document, "WIFI Free");
                }

                /// <summary>
                /// Imprimir el reporte
                /// </summary>
                internal void Print() => Helper.Print(document, @"MP-4000 TH");
            }
            internal class WifiType
            {
                string room;
                string login;
                string pass;
                short freeUsers;

                public string Room { get => room.ToUpper(); set => room = value; }
                public string Login { get => login; set => login = value; }
                public string Pass { get => pass; set => pass = value; }
                public short FreeUsers { get => freeUsers; set => freeUsers = value; }
            }
        }
    }
}
