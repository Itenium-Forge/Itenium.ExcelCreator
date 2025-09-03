using ClosedXML.Excel;
using Itenium.ExcelCreator.Client;
using Itenium.ExcelCreator.WebApi.Helpers;
using Microsoft.AspNetCore.Mvc;

namespace Itenium.ExcelCreator.WebApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelController : ControllerBase, IExcelService
{
    private readonly ILogger<ExcelController> _logger;

    public ExcelController(ILogger<ExcelController> logger)
    {
        _logger = logger;
    }

    [HttpGet]
    public IActionResult Get()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().SetValue(42);
        return wb.Deliver("excelfile.xlsx");
    }
}
