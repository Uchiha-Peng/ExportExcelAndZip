using ExportExcelTest.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.IO.Compression;

namespace ExportExcelTest
{
    /// <summary>
    /// server 的摘要说明
    /// </summary>
    public class server : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            var Type = context.Request.Params["Type"];
            var Num = context.Request.Params["Num"];
            if (!string.IsNullOrWhiteSpace(Type) && Type == "download" && !string.IsNullOrWhiteSpace(Num))
            {
                try
                {
                    List<AnnualReportIndexViewModels> list = Tools.GetData();
                    if (Num == "1")
                    {
                        byte[] byteArray = Tools.ExportExcel(list, "2018");
                        if (byteArray == null || byteArray.Length == 0)
                            throw new Exception("文件不能为空!");
                        context.Response.Clear();
                        context.Response.AddHeader(
                            "Content-Length", byteArray.Length.ToString());
                        context.Response.ContentType = "application/octet-stream";
                        context.Response.AddHeader(
                            "content-disposition",
                            "attachment; filename=医疗质控年报指标.xlsx");
                        context.Response.OutputStream.Write(byteArray, 0, byteArray.Length);
                        context.Response.Flush();
                        context.Response.End();
                    }
                    else
                    {
                        byte[] byteArray = Tools.ExportZip(list, "2018");
                        if (byteArray == null || byteArray.Length == 0)
                            throw new Exception("文件不能为空!");
                        context.Response.Clear();
                        context.Response.AddHeader(
                            "Content-Length", byteArray.Length.ToString());
                        context.Response.ContentType = "application/octet-stream";
                        context.Response.AddHeader(
                            "content-disposition",
                            "attachment; filename=2018医疗质控年报指标.zip");
                        context.Response.OutputStream.Write(byteArray, 0, byteArray.Length);
                        context.Response.Flush();
                        context.Response.End();
                    }
                }
                catch (Exception ex)
                {
                    context.Response.ContentType = "text/html";
                    HttpContext.Current.Response.Write($@"<script>alert('导出Excle出错：" + ex.Message + "')</script><script>");
                }

            }


        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}