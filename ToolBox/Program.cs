using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolBox
{
    class Program
    {
        // Declaración de variables
        static DateTime ReportDateFrom = new DateTime(2018, 04, 01);
        static DateTime ReportDateTo = new DateTime(2018, 04, 30);

        static void Main(string[] args)
        {

            // Solicitud de datos para el reporte
            Console.WriteLine("Reporte de facturación");
            Console.WriteLine("");
            Console.Write($"Fecha inicial ({ReportDateFrom.ToShortDateString()}): ");
            var ReportDateFrom_T = Console.ReadLine();
            if (ReportDateFrom_T != "") ReportDateFrom = DateTime.Parse(ReportDateFrom_T);
            Console.Write($"Fecha final ({ReportDateTo.ToShortDateString()}): ");
            var ReportDateTo_T = Console.ReadLine();
            if (ReportDateTo_T != "") ReportDateTo = DateTime.Parse(ReportDateTo_T);

            // Generacion de los reportes
            GenerateXLS(ReportDateFrom, ReportDateTo);
            Console.ReadLine();
        }
        static void GenerateXLS(DateTime ReportDateFrom, DateTime ReportDateTo)
        {
            // Generación del Libro
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(2);

            //GenerateSheet(workbook.Worksheets[0], ReportDateFrom, ReportDateTo, 1);
            GenerateSheet(workbook.Worksheets[1], ReportDateFrom, ReportDateTo, 2);

            // Guardar el documento
            workbook.SaveAs($"Reporte RIMISOL S.A. {ReportDateFrom.ToString("dd-MM-yy")} - {ReportDateTo.ToString("dd-MM-yy")}.xlsx");
        }

        static void GenerateSheet(IWorksheet worksheet, DateTime ReportDateFrom, DateTime ReportDateTo, int MonId)
        {
            var mon = MonId == 1 ? "UYU" : "USD";
            worksheet.Name = $"Reporte de facturación en {mon}";

            // Generación de los titulos y los datos básicos del reporte
            worksheet.Range[1, 1].Text = "Informe de facturación de Rimisol SA";
            worksheet.Range[2, 1].Text = $"Periodo del reporte: {ReportDateFrom.ToShortDateString()} al {ReportDateTo.ToShortDateString()}";
            worksheet.Range[1, 1, 1, 6].Merge();
            worksheet.Range[2, 1, 2, 6].Merge();

            worksheet.Range[4, 1].Text = "Datos básicos de la factura:";
            worksheet.Range[4, 1, 4, 12].Merge();
            worksheet.Range[5, 1].Text = "Fecha:";
            worksheet.Range[5, 2].Text = "Tipo de CFE:";
            worksheet.Range[5, 2, 5, 3].Merge();
            worksheet.Range[5, 4].Text = "Serie:";
            worksheet.Range[5, 5].Text = "Número:";
            worksheet.Range[5, 6].Text = "Moneda:";
            worksheet.Range[5, 7].Text = "Punto de venta:";
            worksheet.Range[5, 7, 5, 8].Merge();
            worksheet.Range[5, 9].Text = "Nombre:";
            worksheet.Range[5, 9, 5, 10].Merge();
            worksheet.Range[5, 11].Text = "R.U.T.:";
            worksheet.Range[5, 11, 5, 12].Merge();

            worksheet.Range[4, 13].Text = "Hospedaje:";
            worksheet.Range[4, 13, 4, 18].Merge();
            worksheet.Range[5, 13].Text = "No gravado:";
            worksheet.Range[5, 14].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 15].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 16].Text = "IVA Mínimo:";
            worksheet.Range[5, 17].Text = "IVA Básico:";
            worksheet.Range[5, 18].Text = "Total:";

            worksheet.Range[4, 19].Text = "Restaurante:";
            worksheet.Range[4, 19, 4, 24].Merge();
            worksheet.Range[5, 19].Text = "No gravado:";
            worksheet.Range[5, 20].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 21].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 22].Text = "IVA Mínimo:";
            worksheet.Range[5, 23].Text = "IVA Básico:";
            worksheet.Range[5, 24].Text = "Total:";

            worksheet.Range[4, 25].Text = "Minimercado:";
            worksheet.Range[4, 25, 4, 30].Merge();
            worksheet.Range[5, 25].Text = "No gravado:";
            worksheet.Range[5, 26].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 27].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 28].Text = "IVA Mínimo:";
            worksheet.Range[5, 29].Text = "IVA Básico:";
            worksheet.Range[5, 30].Text = "Total:";

            worksheet.Range[4, 31].Text = "Barra:";
            worksheet.Range[4, 31, 4, 36].Merge();
            worksheet.Range[5, 31].Text = "No gravado:";
            worksheet.Range[5, 32].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 33].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 34].Text = "IVA Mínimo:";
            worksheet.Range[5, 35].Text = "IVA Básico:";
            worksheet.Range[5, 36].Text = "Total:";

            worksheet.Range[4, 37].Text = "Lavadero:";
            worksheet.Range[4, 37, 4, 42].Merge();
            worksheet.Range[5, 37].Text = "No gravado:";
            worksheet.Range[5, 38].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 39].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 40].Text = "IVA Mínimo:";
            worksheet.Range[5, 41].Text = "IVA Básico:";
            worksheet.Range[5, 42].Text = "Total:";

            worksheet.Range[4, 43].Text = "Telefono:";
            worksheet.Range[4, 43, 4, 48].Merge();
            worksheet.Range[5, 43].Text = "No gravado:";
            worksheet.Range[5, 44].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 45].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 46].Text = "IVA Mínimo:";
            worksheet.Range[5, 47].Text = "IVA Básico:";
            worksheet.Range[5, 48].Text = "Total:";

            worksheet.Range[4, 49].Text = "Varios:";
            worksheet.Range[4, 49, 4, 54].Merge();
            worksheet.Range[5, 49].Text = "No gravado:";
            worksheet.Range[5, 50].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 51].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 52].Text = "IVA Mínimo:";
            worksheet.Range[5, 53].Text = "IVA Básico:";
            worksheet.Range[5, 54].Text = "Total:";

            worksheet.Range[4, 55].Text = "Eventos:";
            worksheet.Range[4, 55, 4, 60].Merge();
            worksheet.Range[5, 55].Text = "No gravado:";
            worksheet.Range[5, 56].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 57].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 58].Text = "IVA Mínimo:";
            worksheet.Range[5, 59].Text = "IVA Básico:";
            worksheet.Range[5, 60].Text = "Total:";

            worksheet.Range[4, 61].Text = "Totales:";
            worksheet.Range[4, 61, 4, 67].Merge();
            worksheet.Range[5, 61].Text = "No gravado:";
            worksheet.Range[5, 62].Text = "Sub. IVA Mínimo:";
            worksheet.Range[5, 63].Text = "Sub. IVA Básico:";
            worksheet.Range[5, 64].Text = "IVA Mínimo:";
            worksheet.Range[5, 65].Text = "IVA Básico:";
            worksheet.Range[5, 66].Text = "Impuestos:";
            worksheet.Range[5, 67].Text = "Total:";

            // Obtención de datos de reserva
            List<One.Data.FACTURA> Facturas;
            One.Data.ParadiseDataContext db = new One.Data.ParadiseDataContext();
            Facturas = (from x in db.FACTURA where x.FacFec >= ReportDateFrom & x.FacFec <= ReportDateTo & x.FacCFENumero != 0 & x.FacMoneda == MonId orderby x.FacId select x).ToList();

            int lin = 6;
            int sign = 1;

            double
                TOTAL_HOS_NoGrav = 0,
                TOTAL_HOS_SubMin = 0,
                TOTAL_HOS_SubBas = 0,
                TOTAL_HOS_MIN = 0,
                TOTAL_HOS_BAS = 0,
                TOTAL_HOS_TOT = 0,
                TOTAL_RES_NoGrav = 0,
                TOTAL_RES_SubMin = 0,
                TOTAL_RES_SubBas = 0,
                TOTAL_RES_MIN = 0,
                TOTAL_RES_BAS = 0,
                TOTAL_RES_TOT = 0,
                TOTAL_MIN_NoGrav = 0,
                TOTAL_MIN_SubMin = 0,
                TOTAL_MIN_SubBas = 0,
                TOTAL_MIN_MIN = 0,
                TOTAL_MIN_BAS = 0,
                TOTAL_MIN_TOT = 0,
                TOTAL_BAR_NoGrav = 0,
                TOTAL_BAR_SubMin = 0,
                TOTAL_BAR_SubBas = 0,
                TOTAL_BAR_MIN = 0,
                TOTAL_BAR_BAS = 0,
                TOTAL_BAR_TOT = 0,
                TOTAL_LAV_NoGrav = 0,
                TOTAL_LAV_SubMin = 0,
                TOTAL_LAV_SubBas = 0,
                TOTAL_LAV_MIN = 0,
                TOTAL_LAV_BAS = 0,
                TOTAL_LAV_TOT = 0,
                TOTAL_TEL_NoGrav = 0,
                TOTAL_TEL_SubMin = 0,
                TOTAL_TEL_SubBas = 0,
                TOTAL_TEL_MIN = 0,
                TOTAL_TEL_BAS = 0,
                TOTAL_TEL_TOT = 0,
                TOTAL_VAR_NoGrav = 0,
                TOTAL_VAR_SubMin = 0,
                TOTAL_VAR_SubBas = 0,
                TOTAL_VAR_MIN = 0,
                TOTAL_VAR_BAS = 0,
                TOTAL_VAR_TOT = 0,
                TOTAL_EVE_NoGrav = 0,
                TOTAL_EVE_SubMin = 0,
                TOTAL_EVE_SubBas = 0,
                TOTAL_EVE_MIN = 0,
                TOTAL_EVE_BAS = 0,
                TOTAL_EVE_TOT = 0,
                TOTAL_TOT_NoGrav = 0,
                TOTAL_TOT_SubMin = 0,
                TOTAL_TOT_SubBas = 0,
                TOTAL_TOT_MIN = 0,
                TOTAL_TOT_BAS = 0,
                TOTAL_TOT_IMP = 0,
                TOTAL_TOT_TOT = 0;


            foreach (var Factura in Facturas)
            {
                decimal
                    HOS_NoGrav = 0,
                    HOS_SubMin = 0,
                    HOS_SubBas = 0,
                    HOS_MIN = 0,
                    HOS_BAS = 0,
                    HOS_TOT = 0,
                    RES_NoGrav = 0,
                    RES_SubMin = 0,
                    RES_SubBas = 0,
                    RES_MIN = 0,
                    RES_BAS = 0,
                    RES_TOT = 0,
                    MIN_NoGrav = 0,
                    MIN_SubMin = 0,
                    MIN_SubBas = 0,
                    MIN_MIN = 0,
                    MIN_BAS = 0,
                    MIN_TOT = 0,
                    BAR_NoGrav = 0,
                    BAR_SubMin = 0,
                    BAR_SubBas = 0,
                    BAR_MIN = 0,
                    BAR_BAS = 0,
                    BAR_TOT = 0,
                    LAV_NoGrav = 0,
                    LAV_SubMin = 0,
                    LAV_SubBas = 0,
                    LAV_MIN = 0,
                    LAV_BAS = 0,
                    LAV_TOT = 0,
                    TEL_NoGrav = 0,
                    TEL_SubMin = 0,
                    TEL_SubBas = 0,
                    TEL_MIN = 0,
                    TEL_BAS = 0,
                    TEL_TOT = 0,
                    VAR_NoGrav = 0,
                    VAR_SubMin = 0,
                    VAR_SubBas = 0,
                    VAR_MIN = 0,
                    VAR_BAS = 0,
                    VAR_TOT = 0,
                    EVE_NoGrav = 0,
                    EVE_SubMin = 0,
                    EVE_SubBas = 0,
                    EVE_MIN = 0,
                    EVE_BAS = 0,
                    EVE_TOT = 0;

                Console.WriteLine($"Factura Id {Factura.FacId} [{lin - 5} / {Facturas.Count + 1}]");
                sign =
                    Factura.FacCFETipo.Value == 101 ? 1 :
                    Factura.FacCFETipo.Value == 102 ? -1 :
                    Factura.FacCFETipo.Value == 103 ? 1 :
                    Factura.FacCFETipo.Value == 111 ? 1 :
                    Factura.FacCFETipo.Value == 112 ? -1 :
                    Factura.FacCFETipo.Value == 113 ? 1 : 1;
                worksheet.Range[lin, 1].Text = Factura.FacFec.Value.ToShortDateString();
                worksheet.Range[lin, 2].Text =
                    Factura.FacCFETipo.Value == 101 ? "eTicket" :
                    Factura.FacCFETipo.Value == 102 ? "NC eTicket" :
                    Factura.FacCFETipo.Value == 103 ? "ND eTicket" :
                    Factura.FacCFETipo.Value == 111 ? "eFactura" :
                    Factura.FacCFETipo.Value == 112 ? "NC eFactura" :
                    Factura.FacCFETipo.Value == 113 ? "ND eFactura" : "";
                worksheet.Range[lin, 2, lin, 3].Merge();
                worksheet.Range[lin, 4].Text = Factura.FacCFESerie.TrimEnd(' ');
                worksheet.Range[lin, 5].Number = Factura.FacCFENumero.Value;
                worksheet.Range[lin, 6].Text = Factura.FacMoneda == 1 ? "UYU" : "USD";
                worksheet.Range[lin, 7].Text =
                    Factura.FacSucId == "HOS" ? "Hospedaje" :
                    Factura.FacSucId == "RES" ? "Restaurante" :
                    Factura.FacSucId == "MIN" ? "Minimercado" :
                    Factura.FacSucId == "BAR" ? "Barra" : "";
                worksheet.Range[lin, 7, lin, 8].Merge();
                worksheet.Range[lin, 9].Text = Factura.FacPaxNom.TrimEnd(' ');
                worksheet.Range[lin, 9, lin, 10].Merge();
                worksheet.Range[lin, 11].Text = Factura.FacPaxRuc.Value == 0 ? "" : Factura.FacPaxRuc.Value.ToString();
                worksheet.Range[lin, 11, lin, 12].Merge();


                var FacturaDetalle = (from x in db.FACTURA1 where x.FacId == Factura.FacId select x).ToList();

                foreach (var Detalle in FacturaDetalle)
                {
                    One.Data.FACTURALINIMPUESTO DetalleImp;

                    switch (Detalle.FacPtoV)
                    {
                        case "HOS":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                HOS_NoGrav += Detalle.FacTotLin.Value;
                                HOS_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                HOS_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                HOS_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                HOS_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                HOS_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                HOS_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                HOS_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }

                            HOS_NoGrav = HOS_NoGrav * sign;
                            HOS_SubMin = HOS_SubMin * sign;
                            HOS_SubBas = HOS_SubBas * sign;
                            HOS_MIN = HOS_MIN * sign;
                            HOS_BAS = HOS_BAS * sign;
                            HOS_TOT = HOS_TOT * sign;

                            Console.Write(TOTAL_HOS_TOT + " + " + HOS_TOT + " = ");

                            TOTAL_HOS_NoGrav = TOTAL_HOS_NoGrav + Convert.ToDouble(HOS_NoGrav);
                            TOTAL_HOS_SubMin = TOTAL_HOS_SubMin + Convert.ToDouble(HOS_SubMin);
                            TOTAL_HOS_SubBas = TOTAL_HOS_SubBas + Convert.ToDouble(HOS_SubBas);
                            TOTAL_HOS_MIN = TOTAL_HOS_MIN + Convert.ToDouble(HOS_MIN);
                            TOTAL_HOS_BAS = TOTAL_HOS_BAS + Convert.ToDouble(HOS_BAS);
                            TOTAL_HOS_TOT = TOTAL_HOS_TOT + Convert.ToDouble(HOS_TOT);

                            Console.WriteLine(TOTAL_HOS_TOT);

                            worksheet.Range[lin, 13].Number = Convert.ToDouble(HOS_NoGrav);
                            worksheet.Range[lin, 14].Number = Convert.ToDouble(HOS_SubMin);
                            worksheet.Range[lin, 15].Number = Convert.ToDouble(HOS_SubBas);
                            worksheet.Range[lin, 16].Number = Convert.ToDouble(HOS_MIN);
                            worksheet.Range[lin, 17].Number = Convert.ToDouble(HOS_BAS);
                            worksheet.Range[lin, 18].Number = Convert.ToDouble(HOS_TOT);

                            break;
                        case "RES":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                RES_NoGrav += Detalle.FacTotLin.Value;
                                RES_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                RES_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                RES_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                RES_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                RES_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                RES_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                RES_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            RES_NoGrav = RES_NoGrav * sign;
                            RES_SubMin = RES_SubMin * sign;
                            RES_SubBas = RES_SubBas * sign;
                            RES_MIN = RES_MIN * sign;
                            RES_BAS = RES_BAS * sign;
                            RES_TOT = RES_TOT * sign;

                            TOTAL_RES_NoGrav = TOTAL_RES_NoGrav + Convert.ToDouble(RES_NoGrav);
                            TOTAL_RES_SubMin = TOTAL_RES_SubMin + Convert.ToDouble(RES_SubMin);
                            TOTAL_RES_SubBas = TOTAL_RES_SubBas + Convert.ToDouble(RES_SubBas);
                            TOTAL_RES_MIN = TOTAL_RES_MIN + Convert.ToDouble(RES_MIN);
                            TOTAL_RES_BAS = TOTAL_RES_BAS + Convert.ToDouble(RES_BAS);
                            TOTAL_RES_TOT = TOTAL_RES_TOT + Convert.ToDouble(RES_TOT);

                            worksheet.Range[lin, 19].Number = Convert.ToDouble(RES_NoGrav);
                            worksheet.Range[lin, 20].Number = Convert.ToDouble(RES_SubMin);
                            worksheet.Range[lin, 21].Number = Convert.ToDouble(RES_SubBas);
                            worksheet.Range[lin, 22].Number = Convert.ToDouble(RES_MIN);
                            worksheet.Range[lin, 23].Number = Convert.ToDouble(RES_BAS);
                            worksheet.Range[lin, 24].Number = Convert.ToDouble(RES_TOT);

                            break;
                        case "MIN":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                MIN_NoGrav += Detalle.FacTotLin.Value;
                                MIN_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                MIN_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                MIN_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                MIN_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                MIN_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                MIN_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                MIN_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            MIN_NoGrav = MIN_NoGrav * sign;
                            MIN_SubMin = MIN_SubMin * sign;
                            MIN_SubBas = MIN_SubBas * sign;
                            MIN_MIN = MIN_MIN * sign;
                            MIN_BAS = MIN_BAS * sign;
                            MIN_TOT = MIN_TOT * sign;

                            TOTAL_MIN_NoGrav = TOTAL_MIN_NoGrav + Convert.ToDouble(MIN_NoGrav);
                            TOTAL_MIN_SubMin = TOTAL_MIN_SubMin + Convert.ToDouble(MIN_SubMin);
                            TOTAL_MIN_SubBas = TOTAL_MIN_SubBas + Convert.ToDouble(MIN_SubBas);
                            TOTAL_MIN_MIN = TOTAL_MIN_MIN + Convert.ToDouble(MIN_MIN);
                            TOTAL_MIN_BAS = TOTAL_MIN_BAS + Convert.ToDouble(MIN_BAS);
                            TOTAL_MIN_TOT = TOTAL_MIN_TOT + Convert.ToDouble(MIN_TOT);

                            worksheet.Range[lin, 25].Number = Convert.ToDouble(MIN_NoGrav);
                            worksheet.Range[lin, 26].Number = Convert.ToDouble(MIN_SubMin);
                            worksheet.Range[lin, 27].Number = Convert.ToDouble(MIN_SubBas);
                            worksheet.Range[lin, 28].Number = Convert.ToDouble(MIN_MIN);
                            worksheet.Range[lin, 29].Number = Convert.ToDouble(MIN_BAS);
                            worksheet.Range[lin, 30].Number = Convert.ToDouble(MIN_TOT);

                            break;
                        case "BAR":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                BAR_NoGrav += Detalle.FacTotLin.Value;
                                BAR_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                BAR_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                BAR_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                BAR_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                BAR_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                BAR_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                BAR_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            BAR_NoGrav = BAR_NoGrav * sign;
                            BAR_SubMin = BAR_SubMin * sign;
                            BAR_SubBas = BAR_SubBas * sign;
                            BAR_MIN = BAR_MIN * sign;
                            BAR_BAS = BAR_BAS * sign;
                            BAR_TOT = BAR_TOT * sign;

                            TOTAL_BAR_NoGrav = TOTAL_BAR_NoGrav + Convert.ToDouble(BAR_NoGrav);
                            TOTAL_BAR_SubMin = TOTAL_BAR_SubMin + Convert.ToDouble(BAR_SubMin);
                            TOTAL_BAR_SubBas = TOTAL_BAR_SubBas + Convert.ToDouble(BAR_SubBas);
                            TOTAL_BAR_MIN = TOTAL_BAR_MIN + Convert.ToDouble(BAR_MIN);
                            TOTAL_BAR_BAS = TOTAL_BAR_BAS + Convert.ToDouble(BAR_BAS);
                            TOTAL_BAR_TOT = TOTAL_BAR_TOT + Convert.ToDouble(BAR_TOT);

                            worksheet.Range[lin, 31].Number = Convert.ToDouble(BAR_NoGrav);
                            worksheet.Range[lin, 32].Number = Convert.ToDouble(BAR_SubMin);
                            worksheet.Range[lin, 33].Number = Convert.ToDouble(BAR_SubBas);
                            worksheet.Range[lin, 34].Number = Convert.ToDouble(BAR_MIN);
                            worksheet.Range[lin, 35].Number = Convert.ToDouble(BAR_BAS);
                            worksheet.Range[lin, 36].Number = Convert.ToDouble(BAR_TOT);

                            break;
                        case "LAV":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                LAV_NoGrav += Detalle.FacTotLin.Value;
                                LAV_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                LAV_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                LAV_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                LAV_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                LAV_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                LAV_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                LAV_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            LAV_NoGrav = LAV_NoGrav * sign;
                            LAV_SubMin = LAV_SubMin * sign;
                            LAV_SubBas = LAV_SubBas * sign;
                            LAV_MIN = LAV_MIN * sign;
                            LAV_BAS = LAV_BAS * sign;
                            LAV_TOT = LAV_TOT * sign;

                            TOTAL_LAV_NoGrav = TOTAL_LAV_NoGrav + Convert.ToDouble(LAV_NoGrav);
                            TOTAL_LAV_SubMin = TOTAL_LAV_SubMin + Convert.ToDouble(LAV_SubMin);
                            TOTAL_LAV_SubBas = TOTAL_LAV_SubBas + Convert.ToDouble(LAV_SubBas);
                            TOTAL_LAV_MIN = TOTAL_LAV_MIN + Convert.ToDouble(LAV_MIN);
                            TOTAL_LAV_BAS = TOTAL_LAV_BAS + Convert.ToDouble(LAV_BAS);
                            TOTAL_LAV_TOT = TOTAL_LAV_TOT + Convert.ToDouble(LAV_TOT);

                            worksheet.Range[lin, 37].Number = Convert.ToDouble(LAV_NoGrav);
                            worksheet.Range[lin, 38].Number = Convert.ToDouble(LAV_SubMin);
                            worksheet.Range[lin, 39].Number = Convert.ToDouble(LAV_SubBas);
                            worksheet.Range[lin, 40].Number = Convert.ToDouble(LAV_MIN);
                            worksheet.Range[lin, 41].Number = Convert.ToDouble(LAV_BAS);
                            worksheet.Range[lin, 42].Number = Convert.ToDouble(LAV_TOT);

                            break;
                        case "TEL":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                TEL_NoGrav += Detalle.FacTotLin.Value;
                                TEL_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                TEL_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                TEL_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                TEL_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                TEL_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                TEL_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                TEL_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            TEL_NoGrav = TEL_NoGrav * sign;
                            TEL_SubMin = TEL_SubMin * sign;
                            TEL_SubBas = TEL_SubBas * sign;
                            TEL_MIN = TEL_MIN * sign;
                            TEL_BAS = TEL_BAS * sign;
                            TEL_TOT = TEL_TOT * sign;

                            TOTAL_TEL_NoGrav = TOTAL_TEL_NoGrav + Convert.ToDouble(TEL_NoGrav);
                            TOTAL_TEL_SubMin = TOTAL_TEL_SubMin + Convert.ToDouble(TEL_SubMin);
                            TOTAL_TEL_SubBas = TOTAL_TEL_SubBas + Convert.ToDouble(TEL_SubBas);
                            TOTAL_TEL_MIN = TOTAL_TEL_MIN + Convert.ToDouble(TEL_MIN);
                            TOTAL_TEL_BAS = TOTAL_TEL_BAS + Convert.ToDouble(TEL_BAS);
                            TOTAL_TEL_TOT = TOTAL_TEL_TOT + Convert.ToDouble(TEL_TOT);

                            worksheet.Range[lin, 43].Number = Convert.ToDouble(TEL_NoGrav);
                            worksheet.Range[lin, 44].Number = Convert.ToDouble(TEL_SubMin);
                            worksheet.Range[lin, 45].Number = Convert.ToDouble(TEL_SubBas);
                            worksheet.Range[lin, 46].Number = Convert.ToDouble(TEL_MIN);
                            worksheet.Range[lin, 47].Number = Convert.ToDouble(TEL_BAS);
                            worksheet.Range[lin, 48].Number = Convert.ToDouble(TEL_TOT);

                            break;
                        case "VAR":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                VAR_NoGrav += Detalle.FacTotLin.Value;
                                VAR_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                VAR_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                VAR_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                VAR_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                VAR_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                VAR_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                VAR_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            VAR_NoGrav = VAR_NoGrav * sign;
                            VAR_SubMin = VAR_SubMin * sign;
                            VAR_SubBas = VAR_SubBas * sign;
                            VAR_MIN = VAR_MIN * sign;
                            VAR_BAS = VAR_BAS * sign;
                            VAR_TOT = VAR_TOT * sign;

                            TOTAL_VAR_NoGrav = TOTAL_VAR_NoGrav + Convert.ToDouble(VAR_NoGrav);
                            TOTAL_VAR_SubMin = TOTAL_VAR_SubMin + Convert.ToDouble(VAR_SubMin);
                            TOTAL_VAR_SubBas = TOTAL_VAR_SubBas + Convert.ToDouble(VAR_SubBas);
                            TOTAL_VAR_MIN = TOTAL_VAR_MIN + Convert.ToDouble(VAR_MIN);
                            TOTAL_VAR_BAS = TOTAL_VAR_BAS + Convert.ToDouble(VAR_BAS);
                            TOTAL_VAR_TOT = TOTAL_VAR_TOT + Convert.ToDouble(VAR_TOT);

                            worksheet.Range[lin, 49].Number = Convert.ToDouble(VAR_NoGrav);
                            worksheet.Range[lin, 50].Number = Convert.ToDouble(VAR_SubMin);
                            worksheet.Range[lin, 51].Number = Convert.ToDouble(VAR_SubBas);
                            worksheet.Range[lin, 52].Number = Convert.ToDouble(VAR_MIN);
                            worksheet.Range[lin, 53].Number = Convert.ToDouble(VAR_BAS);
                            worksheet.Range[lin, 54].Number = Convert.ToDouble(VAR_TOT);

                            break;
                        case "EVE":
                            DetalleImp = (from x in db.FACTURALINIMPUESTO where x.FacId == Factura.FacId & x.FacLin == Detalle.FacLin select x).FirstOrDefault();
                            if (DetalleImp == default(One.Data.FACTURALINIMPUESTO) || DetalleImp.FacLinImpuestoIvaTip == 3)
                            {
                                EVE_NoGrav += Detalle.FacTotLin.Value;
                                EVE_TOT += Detalle.FacTotLin.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 1)
                            {
                                EVE_SubMin += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                EVE_MIN += DetalleImp.FacLinImporteImpuestoIva.Value;
                                EVE_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            else if (DetalleImp.FacLinImpuestoIvaTip == 2)
                            {
                                EVE_SubBas += DetalleImp.FacLinImporteImpuestoNeto.Value;
                                EVE_BAS += DetalleImp.FacLinImporteImpuestoIva.Value;
                                EVE_TOT += DetalleImp.FacLinImporteTotal.Value;
                            }
                            EVE_NoGrav = EVE_NoGrav * sign;
                            EVE_SubMin = EVE_SubMin * sign;
                            EVE_SubBas = EVE_SubBas * sign;
                            EVE_MIN = EVE_MIN * sign;
                            EVE_BAS = EVE_BAS * sign;
                            EVE_TOT = EVE_TOT * sign;

                            TOTAL_EVE_NoGrav = TOTAL_EVE_NoGrav + Convert.ToDouble(EVE_NoGrav);
                            TOTAL_EVE_SubMin = TOTAL_EVE_SubMin + Convert.ToDouble(EVE_SubMin);
                            TOTAL_EVE_SubBas = TOTAL_EVE_SubBas + Convert.ToDouble(EVE_SubBas);
                            TOTAL_EVE_MIN = TOTAL_EVE_MIN + Convert.ToDouble(EVE_MIN);
                            TOTAL_EVE_BAS = TOTAL_EVE_BAS + Convert.ToDouble(EVE_BAS);
                            TOTAL_EVE_TOT = TOTAL_EVE_TOT + Convert.ToDouble(EVE_TOT);

                            worksheet.Range[lin, 55].Number = Convert.ToDouble(EVE_NoGrav);
                            worksheet.Range[lin, 56].Number = Convert.ToDouble(EVE_SubMin);
                            worksheet.Range[lin, 57].Number = Convert.ToDouble(EVE_SubBas);
                            worksheet.Range[lin, 58].Number = Convert.ToDouble(EVE_MIN);
                            worksheet.Range[lin, 59].Number = Convert.ToDouble(EVE_BAS);
                            worksheet.Range[lin, 60].Number = Convert.ToDouble(EVE_TOT);

                            break;
                    }

                    TOTAL_TOT_NoGrav += Convert.ToDouble(HOS_NoGrav + RES_NoGrav + MIN_NoGrav + BAR_NoGrav + LAV_NoGrav + TEL_NoGrav + VAR_NoGrav + EVE_NoGrav);
                    TOTAL_TOT_SubMin += Convert.ToDouble(HOS_SubMin + RES_SubMin + MIN_SubMin + BAR_SubMin + LAV_SubMin + TEL_SubMin + VAR_SubMin + EVE_SubMin);
                    TOTAL_TOT_SubBas += Convert.ToDouble(HOS_SubBas + RES_SubBas + MIN_SubBas + BAR_SubBas + LAV_SubBas + TEL_SubBas + VAR_SubBas + EVE_SubBas);
                    TOTAL_TOT_MIN += Convert.ToDouble(HOS_MIN + RES_MIN + MIN_MIN + BAR_MIN + LAV_MIN + TEL_MIN + VAR_MIN + EVE_MIN);
                    TOTAL_TOT_BAS += Convert.ToDouble(HOS_BAS + RES_BAS + MIN_BAS + BAR_BAS + LAV_BAS + TEL_BAS + VAR_BAS + EVE_BAS);
                    TOTAL_TOT_IMP += Convert.ToDouble(HOS_MIN + RES_MIN + MIN_MIN + BAR_MIN + LAV_MIN + TEL_MIN + VAR_MIN + EVE_MIN + HOS_BAS + RES_BAS + MIN_BAS + BAR_BAS + LAV_BAS + TEL_BAS + VAR_BAS + EVE_BAS);
                    TOTAL_TOT_TOT += Convert.ToDouble(Factura.FacTotGen * sign);

                    worksheet.Range[lin, 61].Number = Convert.ToDouble(HOS_NoGrav + RES_NoGrav + MIN_NoGrav + BAR_NoGrav + LAV_NoGrav + TEL_NoGrav + VAR_NoGrav + EVE_NoGrav);
                    worksheet.Range[lin, 62].Number = Convert.ToDouble(HOS_SubMin + RES_SubMin + MIN_SubMin + BAR_SubMin + LAV_SubMin + TEL_SubMin + VAR_SubMin + EVE_SubMin);
                    worksheet.Range[lin, 63].Number = Convert.ToDouble(HOS_SubBas + RES_SubBas + MIN_SubBas + BAR_SubBas + LAV_SubBas + TEL_SubBas + VAR_SubBas + EVE_SubBas);
                    worksheet.Range[lin, 64].Number = Convert.ToDouble(HOS_MIN + RES_MIN + MIN_MIN + BAR_MIN + LAV_MIN + TEL_MIN + VAR_MIN + EVE_MIN);
                    worksheet.Range[lin, 65].Number = Convert.ToDouble(HOS_BAS + RES_BAS + MIN_BAS + BAR_BAS + LAV_BAS + TEL_BAS + VAR_BAS + EVE_BAS);
                    worksheet.Range[lin, 66].Number = Convert.ToDouble(HOS_MIN + RES_MIN + MIN_MIN + BAR_MIN + LAV_MIN + TEL_MIN + VAR_MIN + EVE_MIN + HOS_BAS + RES_BAS + MIN_BAS + BAR_BAS + LAV_BAS + TEL_BAS + VAR_BAS + EVE_BAS);
                    worksheet.Range[lin, 67].Number = Convert.ToDouble(Factura.FacTotGen);

                }

                lin++;
            }

            worksheet.Range[lin + 1, 11].Text = "TOTALES";
            worksheet.Range[lin + 1, 11, lin + 2, 12].Merge();

            worksheet.Range[lin + 1, 13].Text = "Hospedaje:";
            worksheet.Range[lin + 1, 13, lin + 1, 18].Merge();
            worksheet.Range[lin + 2, 13].Text = "No gravado:";
            worksheet.Range[lin + 2, 14].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 15].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 16].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 17].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 18].Text = "Total:";

            worksheet.Range[lin + 1, 19].Text = "Restaurante:";
            worksheet.Range[lin + 1, 19, lin + 1, 24].Merge();
            worksheet.Range[lin + 2, 19].Text = "No gravado:";
            worksheet.Range[lin + 2, 20].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 21].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 22].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 23].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 24].Text = "Total:";

            worksheet.Range[lin + 1, 25].Text = "Minimercado:";
            worksheet.Range[lin + 1, 25, lin + 1, 30].Merge();
            worksheet.Range[lin + 2, 25].Text = "No gravado:";
            worksheet.Range[lin + 2, 26].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 27].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 28].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 29].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 30].Text = "Total:";

            worksheet.Range[lin + 1, 31].Text = "Barra:";
            worksheet.Range[lin + 1, 31, lin + 1, 36].Merge();
            worksheet.Range[lin + 2, 31].Text = "No gravado:";
            worksheet.Range[lin + 2, 32].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 33].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 34].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 35].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 36].Text = "Total:";

            worksheet.Range[lin + 1, 37].Text = "Lavadero:";
            worksheet.Range[lin + 1, 37, lin + 1, 42].Merge();
            worksheet.Range[lin + 2, 37].Text = "No gravado:";
            worksheet.Range[lin + 2, 38].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 39].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 40].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 41].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 42].Text = "Total:";

            worksheet.Range[lin + 1, 43].Text = "Telefono:";
            worksheet.Range[lin + 1, 43, lin + 1, 48].Merge();
            worksheet.Range[lin + 2, 43].Text = "No gravado:";
            worksheet.Range[lin + 2, 44].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 45].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 46].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 47].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 48].Text = "Total:";

            worksheet.Range[lin + 1, 49].Text = "Varios:";
            worksheet.Range[lin + 1, 49, lin + 1, 54].Merge();
            worksheet.Range[lin + 2, 49].Text = "No gravado:";
            worksheet.Range[lin + 2, 50].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 51].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 52].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 53].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 54].Text = "Total:";

            worksheet.Range[lin + 1, 55].Text = "Eventos:";
            worksheet.Range[lin + 1, 55, lin + 1, 60].Merge();
            worksheet.Range[lin + 2, 55].Text = "No gravado:";
            worksheet.Range[lin + 2, 56].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 57].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 58].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 59].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 60].Text = "Total:";

            worksheet.Range[lin + 1, 61].Text = "Totales:";
            worksheet.Range[lin + 1, 61, lin + 1, 67].Merge();
            worksheet.Range[lin + 2, 61].Text = "No gravado:";
            worksheet.Range[lin + 2, 62].Text = "Sub. IVA Mínimo:";
            worksheet.Range[lin + 2, 63].Text = "Sub. IVA Básico:";
            worksheet.Range[lin + 2, 64].Text = "IVA Mínimo:";
            worksheet.Range[lin + 2, 65].Text = "IVA Básico:";
            worksheet.Range[lin + 2, 66].Text = "Impuestos:";
            worksheet.Range[lin + 2, 67].Text = "Total:";

            lin += 3;

            worksheet.Range[lin, 13].Number = Convert.ToDouble(TOTAL_HOS_NoGrav);
            worksheet.Range[lin, 14].Number = Convert.ToDouble(TOTAL_HOS_SubMin);
            worksheet.Range[lin, 15].Number = Convert.ToDouble(TOTAL_HOS_SubBas);
            worksheet.Range[lin, 16].Number = Convert.ToDouble(TOTAL_HOS_MIN);
            worksheet.Range[lin, 17].Number = Convert.ToDouble(TOTAL_HOS_BAS);
            worksheet.Range[lin, 18].Number = Convert.ToDouble(TOTAL_HOS_TOT);
            worksheet.Range[lin, 19].Number = Convert.ToDouble(TOTAL_RES_NoGrav);
            worksheet.Range[lin, 20].Number = Convert.ToDouble(TOTAL_RES_SubMin);
            worksheet.Range[lin, 21].Number = Convert.ToDouble(TOTAL_RES_SubBas);
            worksheet.Range[lin, 22].Number = Convert.ToDouble(TOTAL_RES_MIN);
            worksheet.Range[lin, 23].Number = Convert.ToDouble(TOTAL_RES_BAS);
            worksheet.Range[lin, 24].Number = Convert.ToDouble(TOTAL_RES_TOT);
            worksheet.Range[lin, 25].Number = Convert.ToDouble(TOTAL_MIN_NoGrav);
            worksheet.Range[lin, 26].Number = Convert.ToDouble(TOTAL_MIN_SubMin);
            worksheet.Range[lin, 27].Number = Convert.ToDouble(TOTAL_MIN_SubBas);
            worksheet.Range[lin, 28].Number = Convert.ToDouble(TOTAL_MIN_MIN);
            worksheet.Range[lin, 29].Number = Convert.ToDouble(TOTAL_MIN_BAS);
            worksheet.Range[lin, 30].Number = Convert.ToDouble(TOTAL_MIN_TOT);
            worksheet.Range[lin, 31].Number = Convert.ToDouble(TOTAL_BAR_NoGrav);
            worksheet.Range[lin, 32].Number = Convert.ToDouble(TOTAL_BAR_SubMin);
            worksheet.Range[lin, 33].Number = Convert.ToDouble(TOTAL_BAR_SubBas);
            worksheet.Range[lin, 34].Number = Convert.ToDouble(TOTAL_BAR_MIN);
            worksheet.Range[lin, 35].Number = Convert.ToDouble(TOTAL_BAR_BAS);
            worksheet.Range[lin, 36].Number = Convert.ToDouble(TOTAL_BAR_TOT);
            worksheet.Range[lin, 37].Number = Convert.ToDouble(TOTAL_LAV_NoGrav);
            worksheet.Range[lin, 38].Number = Convert.ToDouble(TOTAL_LAV_SubMin);
            worksheet.Range[lin, 39].Number = Convert.ToDouble(TOTAL_LAV_SubBas);
            worksheet.Range[lin, 40].Number = Convert.ToDouble(TOTAL_LAV_MIN);
            worksheet.Range[lin, 41].Number = Convert.ToDouble(TOTAL_LAV_BAS);
            worksheet.Range[lin, 42].Number = Convert.ToDouble(TOTAL_LAV_TOT);
            worksheet.Range[lin, 43].Number = Convert.ToDouble(TOTAL_TEL_NoGrav);
            worksheet.Range[lin, 44].Number = Convert.ToDouble(TOTAL_TEL_SubMin);
            worksheet.Range[lin, 45].Number = Convert.ToDouble(TOTAL_TEL_SubBas);
            worksheet.Range[lin, 46].Number = Convert.ToDouble(TOTAL_TEL_MIN);
            worksheet.Range[lin, 47].Number = Convert.ToDouble(TOTAL_TEL_BAS);
            worksheet.Range[lin, 48].Number = Convert.ToDouble(TOTAL_TEL_TOT);
            worksheet.Range[lin, 49].Number = Convert.ToDouble(TOTAL_VAR_NoGrav);
            worksheet.Range[lin, 50].Number = Convert.ToDouble(TOTAL_VAR_SubMin);
            worksheet.Range[lin, 51].Number = Convert.ToDouble(TOTAL_VAR_SubBas);
            worksheet.Range[lin, 52].Number = Convert.ToDouble(TOTAL_VAR_MIN);
            worksheet.Range[lin, 53].Number = Convert.ToDouble(TOTAL_VAR_BAS);
            worksheet.Range[lin, 54].Number = Convert.ToDouble(TOTAL_VAR_TOT);
            worksheet.Range[lin, 55].Number = Convert.ToDouble(TOTAL_EVE_NoGrav);
            worksheet.Range[lin, 56].Number = Convert.ToDouble(TOTAL_EVE_SubMin);
            worksheet.Range[lin, 57].Number = Convert.ToDouble(TOTAL_EVE_SubBas);
            worksheet.Range[lin, 58].Number = Convert.ToDouble(TOTAL_EVE_MIN);
            worksheet.Range[lin, 59].Number = Convert.ToDouble(TOTAL_EVE_BAS);
            worksheet.Range[lin, 60].Number = Convert.ToDouble(TOTAL_EVE_TOT);
            worksheet.Range[lin, 61].Number = Convert.ToDouble(TOTAL_TOT_NoGrav);
            worksheet.Range[lin, 62].Number = Convert.ToDouble(TOTAL_TOT_SubMin);
            worksheet.Range[lin, 63].Number = Convert.ToDouble(TOTAL_TOT_SubBas);
            worksheet.Range[lin, 64].Number = Convert.ToDouble(TOTAL_TOT_MIN);
            worksheet.Range[lin, 65].Number = Convert.ToDouble(TOTAL_TOT_BAS);
            worksheet.Range[lin, 66].Number = Convert.ToDouble(TOTAL_TOT_IMP);
            worksheet.Range[lin, 67].Number = Convert.ToDouble(TOTAL_TOT_TOT);


        }
        static void GenerateXLS(List<Factura> Facturas)
        {
            // Generación del Libro
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(2);

            // Generación de los titulos y los datos básicos del reporte
            foreach (var worksheet in workbook.Worksheets.ToList())
            {
                var mon = workbook.Worksheets.ToList().IndexOf(worksheet) == 0 ? "UYU" : "USD";
                worksheet.Name = $"Reporte de facturación en {mon}";

                worksheet.Range[1, 1].Text = "Informe de facturación de Rimisol SA";
                worksheet.Range[2, 1].Text = $"Periodo del reporte: {ReportDateFrom.ToShortDateString()} al {ReportDateTo.ToShortDateString()}";
                worksheet.Range[1, 1, 1, 6].Merge();
                worksheet.Range[2, 1, 2, 6].Merge();

                worksheet.Range[4, 1].Text = "Datos básicos de la factura:";
                worksheet.Range[4, 1, 4, 12].Merge();
                worksheet.Range[5, 1].Text = "Fecha:";
                worksheet.Range[5, 2].Text = "Tipo de CFE:";
                worksheet.Range[5, 2, 5, 3].Merge();
                worksheet.Range[5, 4].Text = "Serie:";
                worksheet.Range[5, 5].Text = "Número:";
                worksheet.Range[5, 6].Text = "Moneda:";
                worksheet.Range[5, 7].Text = "Punto de venta:";
                worksheet.Range[5, 7, 5, 8].Merge();
                worksheet.Range[5, 9].Text = "Nombre:";
                worksheet.Range[5, 9, 5, 10].Merge();
                worksheet.Range[5, 11].Text = "R.U.T.:";
                worksheet.Range[5, 11, 5, 12].Merge();

                worksheet.Range[4, 13].Text = "Hospedaje:";
                worksheet.Range[4, 13, 4, 18].Merge();
                worksheet.Range[5, 13].Text = "No gravado:";
                worksheet.Range[5, 14].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 15].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 16].Text = "IVA Mínimo:";
                worksheet.Range[5, 17].Text = "IVA Básico:";
                worksheet.Range[5, 18].Text = "Total:";

                worksheet.Range[4, 19].Text = "Restaurante:";
                worksheet.Range[4, 19, 4, 24].Merge();
                worksheet.Range[5, 19].Text = "No gravado:";
                worksheet.Range[5, 20].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 21].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 22].Text = "IVA Mínimo:";
                worksheet.Range[5, 23].Text = "IVA Básico:";
                worksheet.Range[5, 24].Text = "Total:";

                worksheet.Range[4, 25].Text = "Minimercado:";
                worksheet.Range[4, 25, 4, 30].Merge();
                worksheet.Range[5, 25].Text = "No gravado:";
                worksheet.Range[5, 26].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 27].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 28].Text = "IVA Mínimo:";
                worksheet.Range[5, 29].Text = "IVA Básico:";
                worksheet.Range[5, 30].Text = "Total:";

                worksheet.Range[4, 31].Text = "Barra:";
                worksheet.Range[4, 31, 4, 36].Merge();
                worksheet.Range[5, 31].Text = "No gravado:";
                worksheet.Range[5, 32].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 33].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 34].Text = "IVA Mínimo:";
                worksheet.Range[5, 35].Text = "IVA Básico:";
                worksheet.Range[5, 36].Text = "Total:";

                worksheet.Range[4, 37].Text = "Lavadero:";
                worksheet.Range[4, 37, 4, 42].Merge();
                worksheet.Range[5, 37].Text = "No gravado:";
                worksheet.Range[5, 38].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 39].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 40].Text = "IVA Mínimo:";
                worksheet.Range[5, 41].Text = "IVA Básico:";
                worksheet.Range[5, 42].Text = "Total:";

                worksheet.Range[4, 43].Text = "Telefono:";
                worksheet.Range[4, 43, 4, 48].Merge();
                worksheet.Range[5, 43].Text = "No gravado:";
                worksheet.Range[5, 44].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 45].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 46].Text = "IVA Mínimo:";
                worksheet.Range[5, 47].Text = "IVA Básico:";
                worksheet.Range[5, 48].Text = "Total:";

                worksheet.Range[4, 49].Text = "Varios:";
                worksheet.Range[4, 49, 4, 54].Merge();
                worksheet.Range[5, 49].Text = "No gravado:";
                worksheet.Range[5, 50].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 51].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 52].Text = "IVA Mínimo:";
                worksheet.Range[5, 53].Text = "IVA Básico:";
                worksheet.Range[5, 54].Text = "Total:";

                worksheet.Range[4, 55].Text = "Eventos:";
                worksheet.Range[4, 55, 4, 60].Merge();
                worksheet.Range[5, 55].Text = "No gravado:";
                worksheet.Range[5, 56].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 57].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 58].Text = "IVA Mínimo:";
                worksheet.Range[5, 59].Text = "IVA Básico:";
                worksheet.Range[5, 60].Text = "Total:";

                worksheet.Range[4, 61].Text = "Totales:";
                worksheet.Range[4, 61, 4, 67].Merge();
                worksheet.Range[5, 61].Text = "No gravado:";
                worksheet.Range[5, 62].Text = "Sub. IVA Mínimo:";
                worksheet.Range[5, 63].Text = "Sub. IVA Básico:";
                worksheet.Range[5, 64].Text = "IVA Mínimo:";
                worksheet.Range[5, 65].Text = "IVA Básico:";
                worksheet.Range[5, 66].Text = "Impuestos:";
                worksheet.Range[5, 67].Text = "Total:";
            }

            // Index de linea
            int UYULine = 6;
            int USDLine = 6;

            // Generar cuerpo del informe
            foreach (var Factura in Facturas)
            {
                IWorksheet worksheet = workbook.Worksheets[Factura.Moneda - 1];
                int lin = Factura.Moneda == 1 ? UYULine : USDLine;

                worksheet.Range[lin, 1].Text = Factura.Fecha.ToShortDateString();
                worksheet.Range[lin, 2].Text =
                    Factura.TipoCFE == 101 ? "eTicket" :
                    Factura.TipoCFE == 102 ? "NC eTicket" :
                    Factura.TipoCFE == 103 ? "ND eTicket" :
                    Factura.TipoCFE == 111 ? "eFactura" :
                    Factura.TipoCFE == 112 ? "NC eFactura" :
                    Factura.TipoCFE == 113 ? "ND eFactura" : "";
                worksheet.Range[lin, 2, lin, 3].Merge();
                worksheet.Range[lin, 4].Text = Factura.Serie;
                worksheet.Range[lin, 5].Number = Factura.Numero;
                worksheet.Range[lin, 6].Text = Factura.Moneda == 1 ? "UYU" : "USD";
                worksheet.Range[lin, 7].Text =
                    Factura.PuntoDeVenta == "HOS" ? "Hospedaje" :
                    Factura.PuntoDeVenta == "RES" ? "Restaurante" :
                    Factura.PuntoDeVenta == "MIN" ? "Minimercado" :
                    Factura.PuntoDeVenta == "BAR" ? "Barra" : "";
                worksheet.Range[lin, 7, lin, 8].Merge();
                worksheet.Range[lin, 9].Text = Factura.Nombre;
                worksheet.Range[lin, 9, lin, 10].Merge();
                worksheet.Range[lin, 11].Text = Factura.RUT;
                worksheet.Range[lin, 11, lin, 12].Merge();

                foreach (var Detalle in Factura.Detalles)
                {
                    var NoGravado = Detalle.NoGravado;
                    var SubMin = Detalle.SubMin;
                    var SubBas = Detalle.SubBas;
                    var Min = Detalle.Min;
                    var Bas = Detalle.Bas;
                    var Total = Detalle.Total;

                    switch (Detalle.PuntoDeVenta)
                    {
                        case "HOS":
                            worksheet.Range[lin, 13].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 14].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 15].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 16].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 17].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 18].Number = Convert.ToDouble(Total);
                            break;
                        case "RES":
                            worksheet.Range[lin, 19].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 20].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 21].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 22].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 23].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 24].Number = Convert.ToDouble(Total);
                            break;
                        case "MIN":
                            worksheet.Range[lin, 25].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 26].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 27].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 28].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 29].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 30].Number = Convert.ToDouble(Total);
                            break;
                        case "BAR":
                            worksheet.Range[lin, 31].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 32].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 33].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 34].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 35].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 36].Number = Convert.ToDouble(Total);
                            break;
                        case "LAV":
                            worksheet.Range[lin, 37].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 38].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 39].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 40].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 41].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 42].Number = Convert.ToDouble(Total);
                            break;
                        case "TEL":
                            worksheet.Range[lin, 43].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 44].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 45].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 46].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 47].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 48].Number = Convert.ToDouble(Total);
                            break;
                        case "VAR":
                            worksheet.Range[lin, 49].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 50].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 51].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 52].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 53].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 54].Number = Convert.ToDouble(Total);
                            break;
                        case "EVE":
                            worksheet.Range[lin, 55].Number = Convert.ToDouble(NoGravado);
                            worksheet.Range[lin, 56].Number = Convert.ToDouble(SubMin);
                            worksheet.Range[lin, 57].Number = Convert.ToDouble(SubBas);
                            worksheet.Range[lin, 58].Number = Convert.ToDouble(Min);
                            worksheet.Range[lin, 59].Number = Convert.ToDouble(Bas);
                            worksheet.Range[lin, 60].Number = Convert.ToDouble(Total);
                            break;
                    }


                }

                worksheet.Range[lin, 61].Number = Convert.ToDouble(from x in Factura.Detalles group x by x.NoGravado into y select new { z = y.Sum(x => x.NoGravado) });
                worksheet.Range[lin, 62].Number = Convert.ToDouble(from x in Factura.Detalles group x by x.NoGravado into y select new { z = y.Sum(x => x.SubMin) });
                worksheet.Range[lin, 63].Number = Convert.ToDouble(from x in Factura.Detalles group x by x.NoGravado into y select new { z = y.Sum(x => x.SubBas) });
                worksheet.Range[lin, 64].Number = Convert.ToDouble(from x in Factura.Detalles group x by x.NoGravado into y select new { z = y.Sum(x => x.Min) });
                worksheet.Range[lin, 65].Number = Convert.ToDouble(from x in Factura.Detalles group x by x.NoGravado into y select new { z = y.Sum(x => x.Bas) });
                worksheet.Range[lin, 66].Number = Convert.ToDouble(0);
                worksheet.Range[lin, 67].Number = Convert.ToDouble(from x in Factura.Detalles group x by x.NoGravado into y select new { z = y.Sum(x => x.Total) });

                lin++;
                if (Factura.Moneda == 1) UYULine = lin;
                else USDLine = lin;
            }
        }
    }
    public class Factura
    {
        public DateTime Fecha { get; set; }
        public int TipoCFE { get; set; }
        public string Serie { get; set; }
        public int Numero { get; set; }
        public int Moneda { get; set; }
        public string PuntoDeVenta { get; set; }
        public string Nombre { get; set; }
        public string RUT { get; set; }
        public List<Detalle> Detalles { get; set; }

        public class Detalle
        {
            public string PuntoDeVenta { get; set; }
            public decimal NoGravado { get; set; }
            public decimal SubMin { get; set; }
            public decimal SubBas { get; set; }
            public decimal Min { get; set; }
            public decimal Bas { get; set; }
            public decimal Total { get { return NoGravado + SubMin + SubBas + Min + Bas; } }
        }
    }
}
