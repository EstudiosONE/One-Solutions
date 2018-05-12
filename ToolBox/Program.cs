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
        static void Main(string[] args)
        {
            GenerateXLS();
        }
        static void GenerateXLS()
        {
            DateTime ReportDateFrom = new DateTime(2018, 04, 01);
            DateTime ReportDateTo = new DateTime(2018, 04, 30);

            // Solicitud de datos para el reporte
            Console.WriteLine("Reporte de facturación");
            Console.WriteLine("");
            Console.Write($"Fecha inicial ({ReportDateFrom.ToShortDateString()}): ");
            var ReportDateFrom_T = Console.ReadLine();
            if (ReportDateFrom_T != "") ReportDateFrom = DateTime.Parse(ReportDateFrom_T);
            Console.Write($"Fecha final ({ReportDateTo.ToShortDateString()}): ");
            var ReportDateTo_T = Console.ReadLine();
            if (ReportDateTo_T != "") ReportDateTo = DateTime.Parse(ReportDateTo_T);


            // Generación del Libro
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Reporte de facturación";

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
            Facturas = (from x in db.FACTURA where x.FacFec >= ReportDateFrom & x.FacFec <= ReportDateTo & x.FacCFENumero != 0 orderby x.FacId select x).ToList();

            int lin = 6;
            int sign = 1;
            foreach (var Factura in Facturas)
            {
                Console.Clear();
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

                decimal
                    HOS_NoGrav = 0,
                    HOS_SubMin = 0,
                    HOS_SubBas = 0,
                    HOS_MIN = 0,
                    HOS_BAS = 0,
                    HOS_TOT = 0;
                decimal
                    RES_NoGrav = 0,
                    RES_SubMin = 0,
                    RES_SubBas = 0,
                    RES_MIN = 0,
                    RES_BAS = 0,
                    RES_TOT = 0;
                decimal
                    MIN_NoGrav = 0,
                    MIN_SubMin = 0,
                    MIN_SubBas = 0,
                    MIN_MIN = 0,
                    MIN_BAS = 0,
                    MIN_TOT = 0;
                decimal
                    BAR_NoGrav = 0,
                    BAR_SubMin = 0,
                    BAR_SubBas = 0,
                    BAR_MIN = 0,
                    BAR_BAS = 0,
                    BAR_TOT = 0;
                decimal
                    LAV_NoGrav = 0,
                    LAV_SubMin = 0,
                    LAV_SubBas = 0,
                    LAV_MIN = 0,
                    LAV_BAS = 0,
                    LAV_TOT = 0;
                decimal
                    TEL_NoGrav = 0,
                    TEL_SubMin = 0,
                    TEL_SubBas = 0,
                    TEL_MIN = 0,
                    TEL_BAS = 0,
                    TEL_TOT = 0;
                decimal
                    VAR_NoGrav = 0,
                    VAR_SubMin = 0,
                    VAR_SubBas = 0,
                    VAR_MIN = 0,
                    VAR_BAS = 0,
                    VAR_TOT = 0;
                decimal
                    EVE_NoGrav = 0,
                    EVE_SubMin = 0,
                    EVE_SubBas = 0,
                    EVE_MIN = 0,
                    EVE_BAS = 0,
                    EVE_TOT = 0;
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

                            worksheet.Range[lin, 13].Number = Convert.ToDouble(HOS_NoGrav * sign);
                            worksheet.Range[lin, 14].Number = Convert.ToDouble(HOS_SubMin * sign);
                            worksheet.Range[lin, 15].Number = Convert.ToDouble(HOS_SubBas * sign);
                            worksheet.Range[lin, 16].Number = Convert.ToDouble(HOS_MIN * sign);
                            worksheet.Range[lin, 17].Number = Convert.ToDouble(HOS_BAS * sign);
                            worksheet.Range[lin, 18].Number = Convert.ToDouble(HOS_TOT * sign);

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

                            worksheet.Range[lin, 19].Number = Convert.ToDouble(RES_NoGrav * sign);
                            worksheet.Range[lin, 20].Number = Convert.ToDouble(RES_SubMin * sign);
                            worksheet.Range[lin, 21].Number = Convert.ToDouble(RES_SubBas * sign);
                            worksheet.Range[lin, 22].Number = Convert.ToDouble(RES_MIN * sign);
                            worksheet.Range[lin, 23].Number = Convert.ToDouble(RES_BAS * sign);
                            worksheet.Range[lin, 24].Number = Convert.ToDouble(RES_TOT * sign);

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

                            worksheet.Range[lin, 25].Number = Convert.ToDouble(MIN_NoGrav * sign);
                            worksheet.Range[lin, 26].Number = Convert.ToDouble(MIN_SubMin * sign);
                            worksheet.Range[lin, 27].Number = Convert.ToDouble(MIN_SubBas * sign);
                            worksheet.Range[lin, 28].Number = Convert.ToDouble(MIN_MIN * sign);
                            worksheet.Range[lin, 29].Number = Convert.ToDouble(MIN_BAS * sign);
                            worksheet.Range[lin, 30].Number = Convert.ToDouble(MIN_TOT * sign);

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

                            worksheet.Range[lin, 31].Number = Convert.ToDouble(BAR_NoGrav * sign);
                            worksheet.Range[lin, 32].Number = Convert.ToDouble(BAR_SubMin * sign);
                            worksheet.Range[lin, 33].Number = Convert.ToDouble(BAR_SubBas * sign);
                            worksheet.Range[lin, 34].Number = Convert.ToDouble(BAR_MIN * sign);
                            worksheet.Range[lin, 35].Number = Convert.ToDouble(BAR_BAS * sign);
                            worksheet.Range[lin, 36].Number = Convert.ToDouble(BAR_TOT * sign);

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

                            worksheet.Range[lin, 37].Number = Convert.ToDouble(LAV_NoGrav * sign);
                            worksheet.Range[lin, 38].Number = Convert.ToDouble(LAV_SubMin * sign);
                            worksheet.Range[lin, 39].Number = Convert.ToDouble(LAV_SubBas * sign);
                            worksheet.Range[lin, 40].Number = Convert.ToDouble(LAV_MIN * sign);
                            worksheet.Range[lin, 41].Number = Convert.ToDouble(LAV_BAS * sign);
                            worksheet.Range[lin, 42].Number = Convert.ToDouble(LAV_TOT * sign);

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


                            worksheet.Range[lin, 43].Number = Convert.ToDouble(TEL_NoGrav * sign);
                            worksheet.Range[lin, 44].Number = Convert.ToDouble(TEL_SubMin * sign);
                            worksheet.Range[lin, 45].Number = Convert.ToDouble(TEL_SubBas * sign);
                            worksheet.Range[lin, 46].Number = Convert.ToDouble(TEL_MIN * sign);
                            worksheet.Range[lin, 47].Number = Convert.ToDouble(TEL_BAS * sign);
                            worksheet.Range[lin, 48].Number = Convert.ToDouble(TEL_TOT * sign);

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


                            worksheet.Range[lin, 49].Number = Convert.ToDouble(VAR_NoGrav * sign);
                            worksheet.Range[lin, 50].Number = Convert.ToDouble(VAR_SubMin * sign);
                            worksheet.Range[lin, 51].Number = Convert.ToDouble(VAR_SubBas * sign);
                            worksheet.Range[lin, 52].Number = Convert.ToDouble(VAR_MIN * sign);
                            worksheet.Range[lin, 53].Number = Convert.ToDouble(VAR_BAS * sign);
                            worksheet.Range[lin, 54].Number = Convert.ToDouble(VAR_TOT * sign);

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


                            worksheet.Range[lin, 55].Number = Convert.ToDouble(EVE_NoGrav * sign);
                            worksheet.Range[lin, 56].Number = Convert.ToDouble(EVE_SubMin * sign);
                            worksheet.Range[lin, 57].Number = Convert.ToDouble(EVE_SubBas * sign);
                            worksheet.Range[lin, 58].Number = Convert.ToDouble(EVE_MIN * sign);
                            worksheet.Range[lin, 59].Number = Convert.ToDouble(EVE_BAS * sign);
                            worksheet.Range[lin, 60].Number = Convert.ToDouble(EVE_TOT * sign);

                            break;
                    }
                }

                lin++;
            }

            // Guardar el documento
            workbook.SaveAs("test.xlsx");
        }
    }
}
