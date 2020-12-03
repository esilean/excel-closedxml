using ClosedXML.Excel;
using ExcelCsvApp.Api.Models;
using ExcelCsvApp.Api.Services;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelCsvApp.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IExcelService _excelService;
        public ExcelController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        [HttpGet]
        public async Task<IActionResult> GetExcelTemplate()
        {
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            string fileName = Guid.NewGuid().ToString() + "historico.xlsx";
            try
            {
                var content = await _excelService.GerarArquivoHistorico(historico, "TemplateTable.xlsx", fileName);
                return File(content, contentType, fileName);

            }
            catch (Exception ex)
            {
                return StatusCode(500, ex);
            }
            finally
            {
                if (System.IO.File.Exists(fileName))
                    System.IO.File.Delete(fileName);
            }
        }


        [HttpGet("{id}")]
        public IActionResult GetExcel(int id)
        {

            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            string fileName = Guid.NewGuid().ToString() + "_users.xlsx";
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    IXLWorksheet worksheet = workbook.Worksheets.Add("Users");
                    worksheet.Cell(1, 1).Value = "Id";
                    worksheet.Cell(1, 2).Value = "Username";
                    for (int index = 1; index <= users.Count; index++)
                    {
                        worksheet.Cell(index + 1, 1).Value = users[index - 1].Id;
                        worksheet.Cell(index + 1, 2).Value = users[index - 1].Username;
                        worksheet.Cell(index + 1, 3).Value = id;
                    }
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, contentType, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex);
            }
        }

        private readonly List<StockHistory> historico = new List<StockHistory>
        {
            new StockHistory {
                Date = DateTime.Now,
                NetWorth = 99999,
                Quota = 88888,
                MediaV1 = 1001,
                ClosedV1 = 1002,
                MediaV2 = 2001,
                ClosedV2 = 2002,
                MediaV3 = 3001,
                ClosedV3 = 3002,
                MediaV4 = 4001,
                ClosedV4 = 4002
            },
            new StockHistory {
                Date = DateTime.Now,
                NetWorth = 99999,
                Quota = 88888,
                MediaV1 = 1001,
                ClosedV1 = 1002,
                MediaV2 = 2001,
                ClosedV2 = 2002,
                MediaV3 = 3001,
                ClosedV3 = 3002,
                MediaV4 = 4001,
                ClosedV4 = 4002
            },
            new StockHistory {
                Date = DateTime.Now,
                NetWorth = 99999,
                Quota = 88888,
                MediaV1 = 1001,
                ClosedV1 = 1002,
                MediaV2 = 2001,
                ClosedV2 = 2002,
                MediaV3 = 3001,
                ClosedV3 = 3002,
                MediaV4 = 4001,
                ClosedV4 = 4002
            }
        };

        private readonly List<User> users = new List<User>
        {
            new User { Id = 1, Username = "DoloresAbernathy" },
            new User { Id = 2, Username = "MaeveMillay" },
            new User { Id = 3, Username = "BernardLowe" },
            new User { Id = 4, Username = "ManInBlack" }
        };
    }
}
