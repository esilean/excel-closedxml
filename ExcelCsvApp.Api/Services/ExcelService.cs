using ClosedXML.Excel;
using ExcelCsvApp.Api.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelCsvApp.Api.Services
{
    public class ExcelService : IExcelService
    {

        public Task<byte[]> GerarArquivoHistorico(List<StockHistory> stockHistory, string pathTemplate, string pathFile)
        {
            File.Copy(pathTemplate, pathFile);

            byte[] rel = null;

            using (var workbook = new XLWorkbook(pathFile))
            {
                var worksheet = workbook.Worksheets.Worksheet("HistoricoCotas");

                for (int i = 0; i < stockHistory.Count; i++)
                {
                    var cotacao = stockHistory[i];
                    worksheet.Cell("A" + (5 + i)).Value = cotacao.Date;
                    worksheet.Cell("B" + (5 + i)).Value = cotacao.NetWorth;
                    worksheet.Cell("C" + (5 + i)).Value = cotacao.Quota;

                    worksheet.Cell("D" + (5 + i)).Value = cotacao.MediaV1;
                    worksheet.Cell("E" + (5 + i)).Value = cotacao.ClosedV1;

                    worksheet.Cell("F" + (5 + i)).Value = cotacao.MediaV2;
                    worksheet.Cell("G" + (5 + i)).Value = cotacao.ClosedV2;

                    worksheet.Cell("H" + (5 + i)).Value = cotacao.MediaV3;
                    worksheet.Cell("I" + (5 + i)).Value = cotacao.ClosedV3;

                    worksheet.Cell("J" + (5 + i)).Value = cotacao.MediaV4;
                    worksheet.Cell("K" + (5 + i)).Value = cotacao.ClosedV4;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    rel = stream.ToArray();
                }
            }

            return Task.FromResult(rel);

        }

    }
}
