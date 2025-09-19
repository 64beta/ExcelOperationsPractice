using ExcelOperationsPractice.DTOs;

namespace ExcelOperationsPractice.Services.Interfaces
{
    public interface IExcelService
    {
        List<EmployeeExcelDTO> ReadExcel(IFormFile file);
        MemoryStream ColorSingleToGreen(IFormFile file);
        MemoryStream GenerateTemplateExcel<T>(Dictionary<string, IEnumerable<(string Code, string Name)>> lookups= null,int dataRows = 100);
        List<T> ReadExcelParzival<T>(IFormFile file) where T : new();
    }
}


