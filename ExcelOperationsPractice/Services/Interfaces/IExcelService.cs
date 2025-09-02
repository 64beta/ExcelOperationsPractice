using ExcelOperationsPractice.DTOs;

namespace ExcelOperationsPractice.Services.Interfaces
{
    public interface IExcelService
    {
        List<EmployeeExcelDTO> ReadExcel(IFormFile file);
    }
}
