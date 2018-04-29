using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace One.Services.Paradise.Hotel
{
    internal class Reservations
    {
        internal class SetRatesPeriodType
        {
            public DateTime From { get; set; }
            public DateTime To { get; set; }
            public Dictionary<DayOfWeek, decimal> Days { get; set; }
            public string Code { get; set; }
            public bool AdmitCheckIn { get; set; }
            public bool AdmitCheckOut { get; set; }
        }

        internal static void SetRateInPeriod(SetRatesPeriodType periodData)
        {
            for (DateTime date = periodData.From; date <= periodData.To; date = date.AddDays(1))
            {
                var amount = (from x in periodData.Days where x.Key == date.DayOfWeek select x.Value).FirstOrDefault();
                if (amount != default(decimal))
                {
                    db.one_iGSDataContext db = new db.one_iGSDataContext();
                    var temporalRate = new db.RateDetail()
                    {
                        Code = periodData.Code,
                        Date = date,
                        Amount = amount,
                        AdmitCheckIn = periodData.AdmitCheckIn,
                        AdmitCheckOut = periodData.AdmitCheckOut
                    };

                    var rateDetail = from x in db.RateDetail where x.Code == periodData.Code & x.Date == date select x;
                    if (rateDetail.Count() == 0)
                    {
                        db.RateDetail.InsertOnSubmit(temporalRate);
                    }
                    else
                    {
                        foreach (db.RateDetail rate in rateDetail)
                        {
                            rate.Amount = temporalRate.Amount;
                            rate.AdmitCheckIn = temporalRate.AdmitCheckIn;
                            rate.AdmitCheckOut = temporalRate.AdmitCheckOut;
                        }
                    }

                    db.SubmitChanges();
                }
                else
                {

                }
            }
        }
    }
}
