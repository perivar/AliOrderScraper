using System;
using System.Globalization;

namespace AliOrderScraper
{
    public class AliExpressOrder
    {
        public long OrderId { get; set; }
        public DateTime OrderTime { get; set; }
        public string StoreName { get; set; }
        public string StoreUrl { get; set; }
        public decimal OrderAmount { get; set; }
        public string OrderLines { get; set; }
        public string ContactName { get; set; }
        public string ContactAddress { get; set; }
        public string ContactAddress2 { get; set; }
        public string ContactZipCode { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1:dd.MM.yyyy} {2} {3}", OrderId, OrderTime, OrderAmount.ToString("C", new CultureInfo("en-US")), ContactName);
        }
    }
}
