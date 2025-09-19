using ExcelOperationsPractice.DTOs;
using ExcelOperationsPractice.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExcelOperationsPractice.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class TemplateController : ControllerBase
    {
        private readonly IExcelService _excelService;

        public TemplateController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        [HttpPost]
        public IActionResult UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Lütfen bir excel dosyası yükleyin.");

            var stopwatch1 = Stopwatch.StartNew();

            var employees = _excelService.ReadExcel(file);

            stopwatch1.Stop();
            Console.WriteLine($"Read File {stopwatch1.Elapsed.TotalSeconds}");

            return Ok(employees);
        }
        [HttpPost]
        public IActionResult ValidateExcel(IFormFile file)
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
        }
        [HttpPost]
        public async Task<IActionResult> CreateTemplateParzival()
        {


            var stopwatch1 = Stopwatch.StartNew();

            var lookups = new Dictionary<string, IEnumerable<(string Code, string Name)>>()
{
    {
        "MeasurementCriteriaCode", new List<(string, string)>
        {
            ("100", "City Hospital"),
            ("200", "Central Clinic"),
            ("300", "Private Medical Center")
        }
    },{
        "WarehouseCode", new List<(string, string)>
        {
            ("100", "City Hospital"),
            ("200", "Central Clinic"),
            ("300", "Private Medical Center")
        }
    },
    {
        "DepartmentCode", new List<(string, string)>
        {
            ("10", "Cardiology"),
            ("20", "Neurology"),
            ("30", "Orthopedics")
        }
    }
};
            var newTemplate = _excelService.GenerateTemplateExcel<MedicalExpenseDto>(lookups);

            stopwatch1.Stop();
            Console.WriteLine($"Generate File {stopwatch1.Elapsed.TotalSeconds}");
            return File(newTemplate, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ColoredEmployees.xlsx");
        }

        [HttpPost]
        public async Task<IActionResult> ReadTemplateParzival(IFormFile file)
        {


            var stopwatch1 = Stopwatch.StartNew();


            List<MedicalExpenseDto> result = _excelService.ReadExcelParzival<MedicalExpenseDto>(file);

            stopwatch1.Stop();
            Console.WriteLine($"Read File {stopwatch1.Elapsed.TotalSeconds}");

            return Ok(result);
        }
        [HttpPost]
        public async Task<IActionResult> CreateTemplateParzival2()
        {


            var stopwatch1 = Stopwatch.StartNew();

            
            var newTemplate = _excelService.GenerateTemplateExcel<EmployeeExcelDTO>();

            stopwatch1.Stop();
            Console.WriteLine($"Generate File {stopwatch1.Elapsed.TotalSeconds}");
            return File(newTemplate, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ColoredEmployees.xlsx");
        }
    }
}