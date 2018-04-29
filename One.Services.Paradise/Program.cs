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
            //dbDataContext db = new dbDataContext();
            //var numeradores = from x in db.NUMERADORES where x.NumId == 51 select x;
            //foreach (var item in numeradores)
            //{
            //    Console.WriteLine($"{item.NumCod}");
            //}
            if (args.Count() > 0)
            {
                switch (args[0].ToLower())
                {
                    default: goto QUERY;
                    case "init": goto INIT;
                }
            }

            QUERY:

            var CHK_IN = new DateTime(2018, 04, 20);
            var CHK_OUT = new DateTime(2018, 04, 23);
            verdata:
            Console.Write($"Periodo desde ({CHK_IN.ToShortDateString()}):");
            var _CHK_IN = Console.ReadLine();
            CHK_IN = _CHK_IN == "" ? CHK_IN : DateTime.Parse(_CHK_IN);
            Console.Write($"Periodo hasta ({CHK_OUT.ToShortDateString()}):");
            var _CHK_OUT = Console.ReadLine();
            CHK_OUT = _CHK_OUT == "" ? CHK_OUT : DateTime.Parse(_CHK_OUT);
            var roomsAvailable = new List<HABITACION>();
            var nights = (CHK_OUT - CHK_IN).Days;


            using (var db = new dbDataContext())
            {
                var q_habitaciones = from x in db.HABITACION where x.HabNum != 999 & x.HabNum != 998 select x;
                var habitaciones = q_habitaciones.ToList();

                var a = from x in db.ALMANA where x.AlmFec >= CHK_IN.AddDays(-1) & x.AlmFec < CHK_OUT select x;

                var b = a.ToList();

                var c = from x in b where x.AlmFec == CHK_IN.AddDays(-1) & x.AlmReserva != 0 select x;

                var d = c.ToList();

                foreach (var dn in d)
                {
                    var dn_hab = (from x in db.RESERVA where x.ResNro == dn.AlmReserva select x).FirstOrDefault();
                    if (dn_hab != default(RESERVA) & dn_hab.ResLateCheckOut == 's' | dn_hab.ResLateCheckOut == 'S')
                    {
                        habitaciones.RemoveAll(x => x.HabNum == dn_hab.ResHab);
                    }
                }

                foreach (var room in habitaciones)
                {
                    if ((from x in b where x.HabNum == room.HabNum & x.AlmFec >= CHK_IN & x.AlmEst == 'L' select x).ToList().Count == nights)
                    {
                        roomsAvailable.Add(room);
                    }
                }

            }

            using (var db_one = new db.one_iGSDataContext())
            using (var db_paradise = new dbDataContext())
            {
                Dictionary<Tuple<string, string>, List<HABITACION>> roomsAvailableGruped = new Dictionary<Tuple<string, string>, List<HABITACION>>();

                var rooms = (from x in db_paradise.HABITACION select x).ToList();
                var roomTypes = (from x in db_paradise.TIPHAB select x).ToList();
                var roomCategoryes = (from x in db_paradise.CATHAB select x).ToList();
                var rates = (from x in db_one.Rates select x).ToList();
                var ratesDetail = (from x in db_one.RateDetail where x.Date >= CHK_IN & x.Date < CHK_OUT select x).ToList();
                List<db.Rates> ratesToRemove = new List<db.Rates>();
                foreach (var rate in rates)
                {
                    if ((from x in ratesDetail where x.Code == rate.Code select x).Count() < nights)
                    {
                        ratesToRemove.Add(rate);
                    }
                }
                foreach (var rate in ratesToRemove)
                {
                    rates.Remove(rate);
                }



                foreach (var rtype in roomTypes)
                {
                    foreach (var rcat in roomCategoryes)
                    {
                        var roomDetailAvailable = (from x in roomsAvailable where x.HabTipo == rtype.TiphCod & x.HabCat == rcat.CatHCod select x).ToList();
                        if (roomDetailAvailable.Count > 0 & (from x in rates where x.RoomType == rtype.TiphCod & x.RoomCategory == rcat.CatHCod select x).ToList().Count > 0)
                        {
                            Console.WriteLine($"{(roomDetailAvailable.Count == 1 ? "Queda" : "Quedan")} {roomDetailAvailable.Count} {(roomDetailAvailable.Count == 1 ? "habitación" : "habitaciones")} de {rtype.TiphDes.TrimEnd(' ')} {rcat.CatHDes.TrimEnd(' ')}.");
                            foreach (var rate in (from x in rates where x.RoomType == rtype.TiphCod & x.RoomCategory == rcat.CatHCod select x).ToList())
                            {
                                decimal price = 0;

                                foreach (var rateDetail in (from x in ratesDetail where x.Code == rate.Code select x).ToList())
                                {
                                    price += rateDetail.Amount;
                                }

                                Console.WriteLine($"---- {rate.Description.TrimEnd(' ').PadRight(80, ' ')}: {rate.Currency} {price}");
                                Console.WriteLine("");
                            }
                        }
                        else
                        {
                            if ((from x in rooms where x.HabTipo == rtype.TiphCod & x.HabCat == rcat.CatHCod select x).ToList().Count > 0)
                            {
                                Console.WriteLine($"Habitacion de {rtype.TiphDes.TrimEnd(' ')} {rcat.CatHDes.TrimEnd(' ')} no tiene disponibilidad.");
                                Console.WriteLine("");
                            }
                        }
                    }
                }
            }

            Console.ReadLine();
            goto verdata;

            INIT:
            var periodData = new One.Services.Paradise.Hotel.Reservations.SetRatesPeriodType();
            string amount = "";

            STRAT:


            Console.Write($"Código ({periodData.Code}):");
            var code = Console.ReadLine();
            periodData.Code = code == "" ? periodData.Code : code;

            Console.Write($"Periodo desde ({periodData.From.ToShortDateString()}):");
            var from = Console.ReadLine();
            periodData.From = from == "" ? periodData.From : DateTime.Parse(from);
            Console.Write($"Periodo hasta ({periodData.To.ToShortDateString()}):");
            var to = Console.ReadLine();
            periodData.To = to == "" ? periodData.To : DateTime.Parse(to);

            periodData.Days = new Dictionary<DayOfWeek, decimal>();
            Console.Write("Domingo: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Sunday, Convert.ToDecimal(amount));
            Console.Write("Lunes: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Monday, Convert.ToDecimal(amount));
            Console.Write("Martes: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Tuesday, Convert.ToDecimal(amount));
            Console.Write("Miercoles: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Wednesday, Convert.ToDecimal(amount));
            Console.Write("Jueves: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Thursday, Convert.ToDecimal(amount));
            Console.Write("Viernes: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Friday, Convert.ToDecimal(amount));
            Console.Write("Sabado: ");
            amount = Console.ReadLine();
            if (amount != "") periodData.Days.Add(DayOfWeek.Saturday, Convert.ToDecimal(amount));

            Console.Write("Admite Check In: ");
            var admitCheckIn = Console.ReadLine().ToLower();
            if (admitCheckIn == "s") periodData.AdmitCheckIn = true; else periodData.AdmitCheckIn = false;
            Console.Write("Admite Check Out: ");
            var admitCheckOut = Console.ReadLine().ToLower();
            if (admitCheckOut == "s") periodData.AdmitCheckOut = true; else periodData.AdmitCheckOut = false;


            Hotel.Reservations.SetRateInPeriod(periodData);


            Console.Clear();

            goto STRAT;

        }
    }
}
