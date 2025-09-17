using ExcelOperationsPractice.DTOs;

namespace ExcelOperationsPractice.Services.Interfaces
{
    public interface IExcelService
    {
        List<EmployeeExcelDTO> ReadExcel(IFormFile file);
        MemoryStream ColorSingleToGreen(IFormFile file);
        List<EmployeeExcelDTO> ReadExcelParzival(IFormFile file);
        public IFormFile GetCustomTemplateParzival();
    }
}
