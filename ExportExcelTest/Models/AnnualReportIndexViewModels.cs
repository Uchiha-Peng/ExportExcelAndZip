using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcelTest.Models
{
    /// <summary>
    /// 年报指标
    /// </summary>
    public class AnnualReportIndexViewModels
    {
        /// <summary>
        /// 部门名称
        /// </summary>
        public string DepName { get; set; }

        /// <summary>
        /// 工作负荷
        /// </summary>
        public List<ReportDetail> Workload { get; set; }

        /// <summary>
        /// 治疗质量
        /// </summary>
        public List<ReportDetail> TreatmentQuality { get; set; }

        /// <summary>
        /// 医疗效率
        /// </summary>
        public List<ReportDetail> MedicalEfficiency { get; set; }

        /// <summary>
        /// 费用情况
        /// </summary>
        public List<ReportDetail> CostSituation { get; set; }
    }
}