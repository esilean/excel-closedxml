using ExcelCsvApp.Api.Models;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;


using System.Text;

namespace ExcelCsvApp.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class CsvController : ControllerBase
    {

        private List<User> users = new List<User>
        {
            new User { Id = 1, Username = "DoloresAbernathy" },
            new User { Id = 2, Username = "MaeveMillay" },
            new User { Id = 3, Username = "BernardLowe" },
            new User { Id = 4, Username = "ManInBlack" }
        };

        [HttpGet]
        public IActionResult GetCsv()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Id;Username");
            foreach (var user in users)
            {
                builder.AppendLine($"{user.Id};{user.Username}");
            }

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
        }
    }
}
