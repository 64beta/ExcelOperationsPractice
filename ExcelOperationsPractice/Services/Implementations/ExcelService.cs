using ExcelOperationsPractice.DTOs;
using ExcelOperationsPractice.Services.Interfaces;
using ClosedXML.Excel;

namespace ExcelOperationsPractice.Services.Implementations
{
    public class ExcelService : IExcelService
{
    public List<EmployeeExcelDTO> ReadExcel(IFormFile file)
    {
        var employees = new List<EmployeeExcelDTO>();

        using (var stream = new MemoryStream())
        {
            file.CopyTo(stream);
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1);
                int id = 1;

                foreach (var row in rows)
                {
                    var dto = new EmployeeExcelDTO
                    {
                        Id = id++,
                        Name = row.Cell(1).GetString(),
                        Surname = row.Cell(2).GetString(),
                        Code = row.Cell(3).GetString(),
                        Job = row.Cell(4).GetString(),
                        StatusCode = row.Cell(5).GetString(),
                        MaritalStatus = row.Cell(6).GetString(),

                    };
                    employees.Add(dto);
                }
            }
        }

        return employees;
    }
}
}
