namespace ExcelOperationsPractice.DTOs
{
    public class MedicalExpenseDto
    {
        public string Name { get; set; }
        public string Manufacturer { get; set; }
        public DateTime IncomeDate { get; set; }
        public DateTime ExpenseDate { get; set; }
        public int ExpenseCount { get; set; }
        public decimal ExpenseAmount { get; set; }

        public int DepartmentCode { get; set; }
        public int MeasurementCriteriaCode { get; set; }
        public int WarehouseCode { get; set; }

    }
}

