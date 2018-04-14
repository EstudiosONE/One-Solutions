using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace One.Services.Printer.API.Hotel
{
    [RoutePrefix("API/Hotel/Wifi")]
    public class WifiController : ApiController
    {
        [Route("Print")]
        [HttpGet]
        public void PrintWifi()
        {
            new Reports.Ticket.Hotel.Wifi(new Reports.Ticket.Hotel.WifiType
            {
                Room = "Santo Domingo",
                Login = "1234567890",
                Pass = "1234567890",
                FreeUsers = 2
            }).Print();

        }
    }
}
