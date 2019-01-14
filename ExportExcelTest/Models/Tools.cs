using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Web;

namespace ExportExcelTest.Models
{
    public class Tools
    {

        public static List<AnnualReportIndexViewModels> GetData()
        {
            List<AnnualReportIndexViewModels> list = new List<AnnualReportIndexViewModels>();
            AnnualReportIndexViewModels vm1 = new AnnualReportIndexViewModels()
            {
                DepName = "检验科",
                Workload = new List<ReportDetail> {
                    new ReportDetail(1, "入院人数"    ),
                    new ReportDetail(2, "出院人数"     ),
                    new ReportDetail(3," 转入人次"    ),
                    new ReportDetail(4, "转出人次"      ),
                    new ReportDetail(5, "出院患者手术人次"   ),
                    new ReportDetail(6,"出院患者手术人次(病案)"),
                    new ReportDetail(7," 单病种出院人次"  ),
                    new ReportDetail(8," 远程会诊工作量" ),
                    new ReportDetail(9," 住院危重患者例数"),
                    new ReportDetail(10," 住院开卡数")
              },
                TreatmentQuality = new List<ReportDetail>() {
                    new ReportDetail(1,"住院患者治愈人数"),
                    new ReportDetail(2,"住院患者治愈率"),
                    new ReportDetail(3,"住院患者死亡人数"),
                    new ReportDetail(4,"住院患者死亡率"),
                    new ReportDetail(5,"单病种住院死亡率"),
                    new ReportDetail(6,"手术患者住院死亡率"),
                    new ReportDetail(7,"入院与出院诊断符合率"),
                    new ReportDetail(8,"门诊与入院诊断符合率"),
                    new ReportDetail(9,"住院抢救成功例数"),
                    new ReportDetail(10,"住院抢救成功率"),
                    new ReportDetail(11,"非计划再手术例数"),
                    new ReportDetail(12,"当天再入院人次"),
                    new ReportDetail(13,"当天再入院率"),
                    new ReportDetail(14,"2-10天内再入院人次"),
                    new ReportDetail(15,"2-10天内再入院率"),
                    new ReportDetail(16,"10-31天内再入院人次"),
                    new ReportDetail(17,"10-31天内再入院率"),
                    new ReportDetail(18,"2-31天内再入院人次"),
                    new ReportDetail(19,"2-31天内再入院率"),
                    new ReportDetail(20,"重返ICU人次"),
                    new ReportDetail(21,"重返ICU发生率"),
                    new ReportDetail(22,"出院患者感染发生例数"),
                    new ReportDetail(23,"出院患者感染发生率"),
                    new ReportDetail(24,"不良事件上报例数" )},
                MedicalEfficiency = new List<ReportDetail> {
                    new ReportDetail(1,"出院患者平均住院日		"),
                    new ReportDetail(2,"入病区平均等待时长"),
                    new ReportDetail(3,"手术平均等待时长"),
                    new ReportDetail(4,"床位使用率"),
                    new ReportDetail(5,"床位周转率")
              },
                CostSituation = new List<ReportDetail>{
                    new ReportDetail(1,"出院患者住院费用"),
                    new ReportDetail(2,"出院患者住院药费"),
                    new ReportDetail(3,"出院患者次均费用"),
                    new ReportDetail(4,"出院患者次均药费"),
                    new ReportDetail(5,"出院患者药费占比"),
                    new ReportDetail(6,"出院患者耗材费占比")
               }
            };
            AnnualReportIndexViewModels vm2 = new AnnualReportIndexViewModels()
            {
                DepName = "眼科",
                Workload = new List<ReportDetail> {
                    new ReportDetail(1, "入院人数"    ),
                    new ReportDetail(2, "出院人数"     ),
                    new ReportDetail(3," 转入人次"    ),
                    new ReportDetail(4, "转出人次"      ),
                    new ReportDetail(5, "出院患者手术人次"   ),
                    new ReportDetail(6,"出院患者手术人次(病案)"),
                    new ReportDetail(7," 单病种出院人次"  ),
                    new ReportDetail(8," 远程会诊工作量" ),
                    new ReportDetail(9," 住院危重患者例数"),
                    new ReportDetail(10," 住院开卡数")
              },
                TreatmentQuality = new List<ReportDetail>() {
                    new ReportDetail(1,"住院患者治愈人数"),
                    new ReportDetail(2,"住院患者治愈率"),
                    new ReportDetail(3,"住院患者死亡人数"),
                    new ReportDetail(4,"住院患者死亡率"),
                    new ReportDetail(5,"单病种住院死亡率"),
                    new ReportDetail(6,"手术患者住院死亡率"),
                    new ReportDetail(7,"入院与出院诊断符合率"),
                    new ReportDetail(8,"门诊与入院诊断符合率"),
                    new ReportDetail(9,"住院抢救成功例数"),
                    new ReportDetail(10,"住院抢救成功率"),
                    new ReportDetail(11,"非计划再手术例数"),
                    new ReportDetail(12,"当天再入院人次"),
                    new ReportDetail(13,"当天再入院率"),
                    new ReportDetail(14,"2-10天内再入院人次"),
                    new ReportDetail(15,"2-10天内再入院率"),
                    new ReportDetail(16,"10-31天内再入院人次"),
                    new ReportDetail(17,"10-31天内再入院率"),
                    new ReportDetail(18,"2-31天内再入院人次"),
                    new ReportDetail(19,"2-31天内再入院率"),
                    new ReportDetail(20,"重返ICU人次"),
                    new ReportDetail(21,"重返ICU发生率"),
                    new ReportDetail(22,"出院患者感染发生例数"),
                    new ReportDetail(23,"出院患者感染发生率"),
                    new ReportDetail(24,"不良事件上报例数" )},
                MedicalEfficiency = new List<ReportDetail> {
                    new ReportDetail(1,"出院患者平均住院日		"),
                    new ReportDetail(2,"入病区平均等待时长"),
                    new ReportDetail(3,"手术平均等待时长"),
                    new ReportDetail(4,"床位使用率"),
                    new ReportDetail(5,"床位周转率")
              },
                CostSituation = new List<ReportDetail>{
                    new ReportDetail(1,"出院患者住院费用"),
                    new ReportDetail(2,"出院患者住院药费"),
                    new ReportDetail(3,"出院患者次均费用"),
                    new ReportDetail(4,"出院患者次均药费"),
                    new ReportDetail(5,"出院患者药费占比"),
                    new ReportDetail(6,"出院患者耗材费占比")
               }
            };
            list.Add(vm1);
            list.Add(vm2);
            return list;
        }



        public static byte[] ExportExcel(List<AnnualReportIndexViewModels> dataList, string Year)
        {
            var path = Path.Combine(HttpContext.Current.Server.MapPath("~"), "data.xlsx");
            try
            {

                if (File.Exists(path))
                    File.Delete(path);
                FileInfo newFile = new FileInfo(path);
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    int rowCount = 51;
                    int columnCout = 5;
                    if (dataList == null || dataList.Count == 0)
                        throw new Exception("列表数据不能为空!");
                    foreach (var item in dataList)
                    {
                        int currentRow = 1;
                        if (item.Workload == null || item.Workload.Count != 10)
                            continue;

                        if (item.TreatmentQuality == null || item.TreatmentQuality.Count != 24)
                            continue;

                        if (item.MedicalEfficiency == null || item.MedicalEfficiency.Count != 5)
                            continue;

                        if (item.CostSituation == null || item.CostSituation.Count != 6)
                            continue;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(Year + "年医疗质控年报指标" + item.DepName);
                        worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Row(1).Style.Font.Size = 16;
                        int i = 1;
                        //设置行高
                        while (i <= rowCount)
                        {
                            worksheet.Row(i).Height = i == 1 ? 22.5 : 16.5;
                            int col = 1;
                            while (col <= 5)
                            {
                                worksheet.Cells[i, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[i, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[i, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                worksheet.Cells[i, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                col++;
                            }
                            i++;
                        }
                        i = 1;
                        //设置列宽
                        while (i <= columnCout)
                        {
                            worksheet.Column(i).Width = i == 2 ? 25 : 10;
                            i++;
                        }
                        //第一行
                        worksheet.Cells[currentRow, 1].Value = Year + "年医疗质控年报指标";
                        worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                        currentRow++;
                        //第二行
                        worksheet.Cells[currentRow, 1].Value = "序号";
                        worksheet.Cells[currentRow, 2].Value = "指标名称";
                        worksheet.Cells[currentRow, 3].Value = "数值";
                        worksheet.Cells[currentRow, 4].Value = "去年数值";
                        worksheet.Cells[currentRow, 5].Value = "同比";
                        currentRow++;
                        //第三行
                        worksheet.Cells[currentRow, 1].Value = "一、工作负荷";
                        worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                        currentRow++;
                        foreach (var vm in item.Workload)
                        {
                            worksheet.Cells[currentRow, 1].Value = vm.Sort;
                            worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                            worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                            worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                            worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                            currentRow++;
                        }
                        //第N行
                        worksheet.Cells[currentRow, 1].Value = "二、治疗质量";
                        worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                        currentRow++;

                        foreach (var vm in item.TreatmentQuality)
                        {
                            worksheet.Cells[currentRow, 1].Value = vm.Sort;
                            worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                            worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                            worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                            worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                            currentRow++;
                        }
                        //第N行
                        worksheet.Cells[currentRow, 1].Value = "三、医疗效率";
                        worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                        currentRow++;

                        foreach (var vm in item.MedicalEfficiency)
                        {
                            worksheet.Cells[currentRow, 1].Value = vm.Sort;
                            worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                            worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                            worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                            worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                            currentRow++;
                        }
                        //第N行
                        worksheet.Cells[currentRow, 1].Value = "四、费用情况";
                        worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                        currentRow++;
                        foreach (var vm in item.CostSituation)
                        {
                            worksheet.Cells[currentRow, 1].Value = vm.Sort;
                            worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                            worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                            worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                            worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                            currentRow++;
                        }
                    }
                    return package.GetAsByteArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static byte[] ExportZip(List<AnnualReportIndexViewModels> dataList, string Year)
        {
            if (dataList == null || dataList.Count == 0)
                throw new Exception("列表数据不能为空!");
            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                    {
                        int rowCount = 51;
                        int columnCout = 5;
                        foreach (var item in dataList)
                        {
                            if (item.Workload == null || item.Workload.Count != 10)
                                continue;

                            if (item.TreatmentQuality == null || item.TreatmentQuality.Count != 24)
                                continue;

                            if (item.MedicalEfficiency == null || item.MedicalEfficiency.Count != 5)
                                continue;

                            if (item.CostSituation == null || item.CostSituation.Count != 6)
                                continue;
                            var path = Path.Combine(HttpContext.Current.Server.MapPath("~"), Year + "年医疗质控年报指标_" + item.DepName + ".xlsx");
                            if (File.Exists(path))
                                File.Delete(path);
                            FileInfo newFile = new FileInfo(path);
                            using (ExcelPackage package = new ExcelPackage(newFile))
                            {
                                int currentRow = 1;
                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(Year + "年医疗质控年报指标_" + item.DepName);
                                worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Row(1).Style.Font.Size = 16;
                                int i = 1;
                                //设置行高
                                while (i <= rowCount)
                                {
                                    worksheet.Row(i).Height = i == 1 ? 22.5 : 16.5;
                                    int col = 1;
                                    while (col <= 5)
                                    {
                                        worksheet.Cells[i, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[i, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[i, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                        worksheet.Cells[i, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                        col++;
                                    }
                                    i++;
                                }
                                i = 1;
                                //设置列宽
                                while (i <= columnCout)
                                {
                                    worksheet.Column(i).Width = i == 2 ? 25 : 10;
                                    i++;
                                }
                                //第一行
                                worksheet.Cells[currentRow, 1].Value = Year + "年医疗质控年报指标";
                                worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                                currentRow++;
                                //第二行
                                worksheet.Cells[currentRow, 1].Value = "序号";
                                worksheet.Cells[currentRow, 2].Value = "指标名称";
                                worksheet.Cells[currentRow, 3].Value = "数值";
                                worksheet.Cells[currentRow, 4].Value = "去年数值";
                                worksheet.Cells[currentRow, 5].Value = "同比";
                                currentRow++;
                                //第三行
                                worksheet.Cells[currentRow, 1].Value = "一、工作负荷";
                                worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                                currentRow++;
                                foreach (var vm in item.Workload)
                                {
                                    worksheet.Cells[currentRow, 1].Value = vm.Sort;
                                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                                    worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                                    worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                                    worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                                    currentRow++;
                                }
                                //第N行
                                worksheet.Cells[currentRow, 1].Value = "二、治疗质量";
                                worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                                currentRow++;

                                foreach (var vm in item.TreatmentQuality)
                                {
                                    worksheet.Cells[currentRow, 1].Value = vm.Sort;
                                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                                    worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                                    worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                                    worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                                    currentRow++;
                                }
                                //第N行
                                worksheet.Cells[currentRow, 1].Value = "三、医疗效率";
                                worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                                currentRow++;

                                foreach (var vm in item.MedicalEfficiency)
                                {
                                    worksheet.Cells[currentRow, 1].Value = vm.Sort;
                                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                                    worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                                    worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                                    worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                                    currentRow++;
                                }
                                //第N行
                                worksheet.Cells[currentRow, 1].Value = "四、费用情况";
                                worksheet.Cells[currentRow, 1, currentRow, columnCout].Merge = true;
                                currentRow++;

                                foreach (var vm in item.CostSituation)
                                {
                                    worksheet.Cells[currentRow, 1].Value = vm.Sort;
                                    worksheet.Cells[currentRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[currentRow, 2].Value = vm.ReportName;
                                    worksheet.Cells[currentRow, 3].Value = vm.CurrentValue;
                                    worksheet.Cells[currentRow, 4].Value = vm.PastValue;
                                    worksheet.Cells[currentRow, 5].Value = vm.ContrastValue;
                                    currentRow++;
                                }
                                ZipArchiveEntry zip = archive.CreateEntry(Year + "年医疗质控年报指标_" + item.DepName + ".xlsx");
                                using (var entryStream = zip.Open())
                                {
                                    byte[] array = package.GetAsByteArray();
                                    entryStream.Write(array, 0, array.Length);
                                }
                            }
                        }

                    }
                    return memoryStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}