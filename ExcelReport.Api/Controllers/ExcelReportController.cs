using ClosedXML.Report;
using ExcelReport.Api.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelReport.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelReportController : ControllerBase
    {

        [HttpGet]
        public IActionResult GetExcelReport()
        {
            string outputFile = $"{Guid.NewGuid()}_HistoricoCotas.xlsx";

            try
            {
                using (var template = new XLTemplate("TemplateTable.xlsx"))
                {
                    template.AddVariable(new StockHeader { StockHists = GetHist() });
                    template.Generate();
                    template.SaveAs(outputFile);
                }

                //Show report
                Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });

                return Ok();
            }
            catch (Exception)
            {
                return StatusCode(500);
            }

        }

        private List<StockHistory> GetHist()
        {
            var hist = new List<StockHistory>();

            for (var i = 1; i <= 1000; i++)
            {
                hist.Add(
                        new StockHistory
                        {
                            Date = DateTime.Now.AddDays(i),
                            NetWorth = i * 2.5,
                            Quota = i * 3.5,
                            MediaV1 = i * 4.5,
                            ClosedV1 = i * 5.5,
                            MediaV2 = i * 6.5,
                            ClosedV2 = i * 7.5,
                            MediaV3 = i * 8.5,
                            ClosedV3 = i * 9.5,
                            MediaV4 = i * 10.5,
                            ClosedV4 = i * 11.5
                        }
                    );
            }

            return hist;

        }
    }
}
