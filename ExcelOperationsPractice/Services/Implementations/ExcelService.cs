using ExcelOperationsPractice.DTOs;
using ExcelOperationsPractice.Services.Interfaces;
using ClosedXML.Excel;

namespace ExcelOperationsPractice.Services.Implementations
{
    public class ExcelService : IExcelService
    {
        private void ColorCells(IXLWorksheet worksheet, IEnumerable<string> cellAddresses, XLColor color)
        {
            foreach (var address in cellAddresses)
            {
                var cell = worksheet.Cell(address);
                cell.Style.Fill.BackgroundColor = color;
            }
        }

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
        public MemoryStream ColorSingleToGreen(IFormFile file)
        {
            var memoryStream = new MemoryStream();
            file.CopyTo(memoryStream);
            memoryStream.Position = 0;

            using (var workbook = new XLWorkbook(memoryStream))
            {
                var worksheet = workbook.Worksheets.First();
                var usedRows = worksheet.RangeUsed().RowsUsed().Skip(1);
                var singleCells = new List<string>();

                foreach (var row in usedRows)
                {
                    var cell = row.Cell(6); 
                    if (cell.GetString().Trim().Equals("Single", StringComparison.OrdinalIgnoreCase))
                    {
                        singleCells.Add(cell.Address.ToString());
                    }
                }

                ColorCells(worksheet, singleCells, XLColor.Green);

                var outputStream = new MemoryStream();
                workbook.SaveAs(outputStream);
                outputStream.Position = 0;
                return outputStream;
            }
        }

    }
}
