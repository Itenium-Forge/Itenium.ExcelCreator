using Microsoft.AspNetCore.Mvc;

namespace Itenium.ExcelCreator.WebApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public string Get()
        {
            return "TheExcel";
        }
    }
}
