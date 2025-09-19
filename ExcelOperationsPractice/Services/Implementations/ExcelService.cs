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
        public MemoryStream GenerateTemplateExcel<T>(
            Dictionary<string, IEnumerable<(string Code, string Name)>> lookups = null,
            int dataRows = 100
        )
        {
            lookups ??= new Dictionary<string, IEnumerable<(string Code, string Name)>>();

            var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Main");

            var props = typeof(T).GetProperties();
            int col = 1;

            foreach (var prop in props)
            {
                if (prop.Name.EndsWith("Code") &&
                    lookups.ContainsKey(prop.Name) &&
                    lookups[prop.Name] != null)
                {
                    var baseProp = prop.Name.Replace("Code", "");
                    var baseName = prop.Name.Replace("Code", "Name");

                    sheet.Cell(1, col).Value = baseName;
                    sheet.Cell(1, col + 1).Value = prop.Name;

                    for (int headerCol = col; headerCol <= col + 1; headerCol++)
                    {
                        var headerCell = sheet.Cell(1, headerCol);
                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Font.FontSize = 11;
                        headerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        headerCell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        headerCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        headerCell.Style.Border.OutsideBorderColor = XLColor.Black;
                    }

                    var lookupSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == baseProp)
                                      ?? workbook.AddWorksheet(baseProp);

                    lookupSheet.Cell(1, 1).Value = "Name";
                    lookupSheet.Cell(1, 2).Value = "Code";
                    for (int headerCol = 1; headerCol <= 2; headerCol++)
                    {
                        var headerCell = lookupSheet.Cell(1, headerCol);
                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        headerCell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        headerCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        headerCell.Style.Border.OutsideBorderColor = XLColor.Black;
                    }

                    var data = lookups[prop.Name].ToList();
                    int startRow = 2;

                    for (int i = 0; i < data.Count; i++)
                    {
                        lookupSheet.Cell(startRow + i, 1).Value = data[i].Name;
                        lookupSheet.Cell(startRow + i, 2).Value = data[i].Code;
                    }

                    var lastRow = startRow + data.Count - 1;

                    if (data.Any())
                    {
                        var nameRange = sheet.Range(2, col, dataRows + 1, col);
                        var lookupRange = lookupSheet.Range(startRow, 1, lastRow, 1);
                        nameRange.SetDataValidation().List(lookupRange);
                    }

                    for (int row = 2; row <= dataRows + 1; row++)
                    {
                        var nameCell = sheet.Cell(row, col);
                        var codeCell = sheet.Cell(row, col + 1);
                        codeCell.FormulaA1 =
                            $"=IFERROR(VLOOKUP({nameCell.Address}, '{baseProp}'!$A:$B, 2, FALSE), \"\")";
                    }

                    col += 2;
                }
                else
                {
                    sheet.Cell(1, col).Value = prop.Name;
                    sheet.Cell(1, col).Style.Font.Bold = true;
                    sheet.Cell(1, col).Style.Fill.BackgroundColor = XLColor.LightGray;
                    sheet.Cell(1, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell(1, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    if (prop.PropertyType == typeof(DateTime) || prop.PropertyType == typeof(DateTime?))
                        sheet.Column(col).Style.DateFormat.Format = "yyyy-MM-dd";
                    else if (prop.PropertyType == typeof(int) || prop.PropertyType == typeof(int?))
                        sheet.Column(col).Style.NumberFormat.Format = "0";
                    else if (prop.PropertyType == typeof(decimal) || prop.PropertyType == typeof(decimal?))
                        sheet.Column(col).Style.NumberFormat.Format = "#,##0.00";

                    col++;
                }
            }

            sheet.Columns().AdjustToContents();
            foreach (var ws in workbook.Worksheets)
                ws.Columns().AdjustToContents();

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            return stream;
        }





        public List<T> ReadExcelParzival<T>(IFormFile file) where T : new()
        {
            var result = new List<T>();

            using var stream = new MemoryStream();
            file.CopyTo(stream);
            stream.Position = 0;

            using var workbook = new XLWorkbook(stream);
            var sheet = workbook.Worksheet("Main");
            if (sheet == null)
                throw new Exception("Excel dosyasında 'Main' isimli sheet bulunamadı.");

            var headers = new Dictionary<int, string>();
            int lastCol = sheet.LastColumnUsed().ColumnNumber();
            for (int col = 1; col <= lastCol; col++)
            {
                var header = sheet.Cell(1, col).GetString();
                if (!string.IsNullOrWhiteSpace(header))
                    headers[col] = header;
            }

            int lastRow = sheet.LastRowUsed().RowNumber();

            for (int row = 2; row <= lastRow; row++)
            {
                var dto = new T();
                bool hasValue = false;

                foreach (var kvp in headers)
                {
                    int col = kvp.Key;
                    string header = kvp.Value;
                    var prop = typeof(T).GetProperty(header);
                    if (prop == null) continue;

                    var cell = sheet.Cell(row, col);
                    if (cell.IsEmpty()) continue;

                    try
                    {
                        object? convertedValue = null;

                        if (prop.PropertyType == typeof(string))
                        {
                            convertedValue = cell.GetString();
                        }
                        else if (prop.PropertyType == typeof(int) || prop.PropertyType == typeof(int?))
                        {
                            convertedValue = cell.TryGetValue<int>(out var intVal) ? intVal : (int?)null;
                        }
                        else if (prop.PropertyType == typeof(decimal) || prop.PropertyType == typeof(decimal?))
                        {
                            convertedValue = cell.TryGetValue<decimal>(out var decVal) ? decVal : (decimal?)null;
                        }
                        else if (prop.PropertyType == typeof(DateTime) || prop.PropertyType == typeof(DateTime?))
                        {
                            if (cell.DataType == XLDataType.DateTime)
                            {
                                convertedValue = cell.GetDateTime();
                            }
                            else if (DateTime.TryParse(cell.GetString(), out var dateVal))
                            {
                                convertedValue = dateVal;
                            }
                            else
                            {
                                convertedValue = prop.PropertyType == typeof(DateTime) ? DateTime.MinValue : (DateTime?)null;
                            }
                        }
                        else
                        {
                            convertedValue = cell.Value.ToString();
                        }

                        if (convertedValue != null && !(convertedValue is string s && string.IsNullOrWhiteSpace(s)))
                            hasValue = true;

                        prop.SetValue(dto, convertedValue);
                    }
                    catch
                    {
                    }
                }

                if (hasValue)
                    result.Add(dto);
            }

            return result;
        }



    }

}

