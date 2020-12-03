using System;

namespace ExcelCsvApp.Api.Models
{
    public class StockHistory
    {
        public DateTime Date { get; set; }
        public double NetWorth { get; set; }
        public double Quota { get; set; }
        public double MediaV1 { get; set; }
        public double ClosedV1 { get; set; }
        public double MediaV2 { get; set; }
        public double ClosedV2 { get; set; }
        public double MediaV3 { get; set; }
        public double ClosedV3 { get; set; }
        public double MediaV4 { get; set; }
        public double ClosedV4 { get; set; }
    }
}
