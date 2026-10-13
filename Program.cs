
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        Excel.Application excelApp = null;
        Excel.Workbook workbook = null;
        Excel.Worksheet worksheet = null;

        try
        {
            // Get the Excel file path
            Console.Write("Enter the full Excel file path: ");
            string filePath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(filePath) ||
                !System.IO.File.Exists(filePath))
            {
                Console.WriteLine("Excel file not found.");
                return;
            }

            // Start Excel
            excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            // Open the workbook
            workbook = excelApp.Workbooks.Open(filePath);

            // Select the worksheet named Data
            worksheet = (Excel.Worksheet)workbook.Worksheets["Data"];

            // Delete the second physical row
            Excel.Range secondRow =
                (Excel.Range)worksheet.Rows[2];

            secondRow.Delete(
                Excel.XlDeleteShiftDirection.xlShiftUp);

            // Save and close
            workbook.Save();
            workbook.Close(false);

            Console.WriteLine(
                "Second row deleted successfully. Excel file saved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            if (worksheet != null)
                Marshal.ReleaseComObject(worksheet);

            if (workbook != null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        Console.WriteLine("Press Enter to exit.");
        Console.ReadLine();
    }
}
