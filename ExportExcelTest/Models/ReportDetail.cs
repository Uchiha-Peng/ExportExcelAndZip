using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcelTest.Models
{
    public class ReportDetail
    {

        public ReportDetail(int sort, string repName)
        {
            Random random = new Random();
            decimal current = random.Next(50, 100);
            decimal past = random.Next(50, 100);
            this.Sort = sort;
            this.ReportName = repName;
            this.CurrentValue = current;
            this.PastValue = past;

        }
        public string ReportName { get; set; }

        public decimal CurrentValue { get; set; }

        public decimal PastValue { get; set; }

        public decimal ContrastValue
        {
            get
            {
                if (CurrentValue > 0 && PastValue > 0)
                {
                    return CurrentValue / PastValue;
                }
                else
                {
                    return 0;
                }
            }
        }

        public int Sort { get; set; }

    }
}