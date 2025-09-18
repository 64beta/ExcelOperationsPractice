using ExcelOperationsPractice.DTOs;

namespace ExcelOperationsPractice.Services.Interfaces
{
    public interface IExcelService
    {
        List<EmployeeExcelDTO> ReadExcel(IFormFile file);
        MemoryStream ColorSingleToGreen(IFormFile file);
        List<EmployeeExcelDTO> ReadExcelParzival(IFormFile file);
        MemoryStream GenerateTemplateExcel<T>(
            Dictionary<string, IEnumerable<(string Code, string Name)>> lookups,
            int dataRows = 100
        );
    }
}
