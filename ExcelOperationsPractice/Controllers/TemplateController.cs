using ExcelOperationsPractice.Services.Interfaces;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelOperationsPractice.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TemplateController : ControllerBase
    {
        private readonly IExcelService _excelService;

        public TemplateController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        [HttpPost("upload")]
        public IActionResult UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Lütfen bir excel dosyası yükleyin.");

            var employees = _excelService.ReadExcel(file);

            return Ok(employees); 
        }
    }
}
