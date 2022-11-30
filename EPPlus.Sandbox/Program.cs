using System;
using System.IO;
using OfficeOpenXml;

namespace EPPlus.Sandbox {
    internal class Program {
        private static void Main(string[] args) {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //var excelFile = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FormulaCalcSample.xlsx"));
            var excelFile = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CalcSheet.xlsx"));
            if (!excelFile.Exists) {
                Console.WriteLine($"Didn't find a file at {excelFile.FullName}");
                Console.ReadLine();
                Environment.Exit(0);
            }


            Console.WriteLine($"Using {excelFile.FullName}");
            Console.WriteLine();
            Console.WriteLine(" ### USER INPUTS ###");
            Console.WriteLine();
            Console.Write("Value 1:");
            var v1 = double.Parse(Console.ReadLine());
            Console.Write("Value 2:");
            var v2 = double.Parse(Console.ReadLine());
            Console.WriteLine();


            using (var package = new ExcelPackage(excelFile)) {
                var dataWorksheet = package.Workbook.Worksheets["Data"];
                var calcWorksheet = package.Workbook.Worksheets["Calculation"];
                dataWorksheet.Cells["B1"].Value = v1;
                dataWorksheet.Cells["B2"].Value = v2;

                RemoveCalculatedFormulaValuesInWorksheet(calcWorksheet);
                //calcWorksheet.Calculate(x => x.AllowCircularReferences = true);
                dataWorksheet.Cells["B3"].Calculate(x => x.AllowCircularReferences = true);
                
                Console.WriteLine(" ### CALCULATION RESULTS ### ");
                Console.WriteLine();
                Console.WriteLine($"Data B1: Value 1 = {dataWorksheet.Cells["B1"].Value}");
                Console.WriteLine($"Data B2: Value 2 = {dataWorksheet.Cells["B2"].Value}");
                Console.WriteLine($"Calc1: {calcWorksheet.Cells["B1"].Formula} = {calcWorksheet.Cells["B1"].Value}");
                Console.WriteLine($"Calc2: {calcWorksheet.Cells["B2"].Formula} = {calcWorksheet.Cells["B2"].Value}");
                Console.WriteLine($"Result: {calcWorksheet.Cells["B3"].Formula} = {calcWorksheet.Cells["B3"].Value}");
                Console.WriteLine(
                    $"Recopied to Data: {dataWorksheet.Cells["B3"].Formula} = {dataWorksheet.Cells["B3"].Value}");

                Console.ReadLine();
            }
        }

        private static void RemoveCalculatedFormulaValuesInWorkbook(ExcelWorkbook workbook) {
            foreach (var worksheet in workbook.Worksheets)
            foreach (var cell in worksheet.Cells)
                // if there is a formula in the cell, the following code keeps the formula but clears the calculated value.
                if (!string.IsNullOrEmpty(cell.Formula)) {
                    var formula = cell.Formula;
                    cell.Value = null;
                    cell.Formula = formula;
                }
        }

        private static void RemoveCalculatedFormulaValuesInWorksheet(ExcelWorksheet worksheet) {
            foreach (var cell in worksheet.Cells)
                // if there is a formula in the cell, the following code keeps the formula but clears the calculated value.
                if (!string.IsNullOrEmpty(cell.Formula)) {
                    var formula = cell.Formula;
                    cell.Value = null;
                    cell.Formula = formula;
                }
        }
    }
}