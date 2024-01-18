using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JSE_Reports.Definition
{
    public class Domain
    {
        public class DailyMTM
        {
            public DateTime FileDate { get; set; }
            public string Contract { get; set; }
            public DateTime ExpiryDate { get; set; }
            public string Classification { get; set; }
            public double Strike { get; set; }
            public string CallPut { get; set; }
            public double MTMYield { get; set; }
            public double MarkPrice { get; set; }
            public double SpotRate { get; set; }
            public double PreviousMTM { get; set; }
            public double PreviousPrice { get; set; }
            public double PremiumOnOption { get; set; }
            public double Volatility { get; set; }
            public double Delta { get; set; }
            public double DeltaValue { get; set; }
            public double ContractsTraded { get; set; }
            public double OpenInterest { get; set; }
        }
    }
}