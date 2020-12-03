using ExcelCsvApp.Api.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelCsvApp.Api.Services
{
    public interface IExcelService
    {
        Task<byte[]> GerarArquivoHistorico(List<StockHistory> stockHistory, string pathTemplate, string path);
    }
}
