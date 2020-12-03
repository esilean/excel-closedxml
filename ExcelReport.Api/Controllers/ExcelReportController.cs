using ClosedXML.Report;
using ExcelReport.Api.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelReport.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelReportController : ControllerBase
    {
        private readonly ILogger<ExcelReportController> _logger;

        public ExcelReportController(ILogger<ExcelReportController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IActionResult GetExcelReport()
        {
            string outputFile = $"{Guid.NewGuid()}_HistoricoCotas.xlsx";

            try
            {
                _logger.LogInformation("Iniciando...");

                using (var template = new XLTemplate("TemplateTable.xlsx"))
                {
                    template.AddVariable(new StockHeader { StockHists = GetHist() });
                    template.Generate();
                    template.SaveAs(outputFile);
                }

                _logger.LogInformation("Processando...");

                //Show report
                Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });

                _logger.LogInformation("Finalizando...");

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, ex.Message);

                return StatusCode(500, ex);
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
