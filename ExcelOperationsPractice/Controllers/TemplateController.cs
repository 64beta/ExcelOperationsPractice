using ExcelOperationsPractice.Services.Interfaces;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

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

            var stopwatch1 = Stopwatch.StartNew();

            var employees = _excelService.ReadExcel(file);

            stopwatch1.Stop();
            Console.WriteLine($"Read File {stopwatch1.Elapsed.TotalSeconds}");





            var stopwatch2 = Stopwatch.StartNew();
            var coloredStream = _excelService.ColorSingleToGreen(file);
            stopwatch2.Stop();
            Console.WriteLine($"ColorSingleToGreen {stopwatch2.Elapsed.TotalSeconds}");
            
            return File(coloredStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ColoredEmployees.xlsx");
            //return Ok(employees); 
        }
    }
}
