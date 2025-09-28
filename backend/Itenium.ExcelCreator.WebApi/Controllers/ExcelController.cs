using Itenium.ExcelCreator.WebApi.Helpers;
using Microsoft.AspNetCore.Mvc;
using Itenium.ExcelCreator.Library;
using Itenium.ExcelCreator.Library.Models;

namespace Itenium.ExcelCreator.WebApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelController(ExcelService service) : ControllerBase
{
    /// <summary>
    /// Creates an Excel from the data and configuration posted in the body
    /// </summary>
    /// <returns>The generated Excel</returns>
    [HttpPost]
    public FileStreamResult Create([FromBody] FullExcelData data)
    {
        var wb = service.CreateExcel(data);
        return wb.Deliver(data.Config.FileName);
    }
}
