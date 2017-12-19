using System;

namespace AliOrderScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentDate = DateTime.Now.Date;
            var currentYear = currentDate.Year;
            var from = new DateTime(currentYear, 1, 1);
            var to = currentDate;

            AliExpress.ScrapeAliExpressOrders("", "AliExpressOrders", from);
        }
    }
}
