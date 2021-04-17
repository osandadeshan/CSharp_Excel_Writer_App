using GemBox.Spreadsheet;

namespace ExcelWriterApp.Util
{
    public static class ExcelWriter
    {
        public static void WriteAsNewExcelFile(string excelFilePath, string sheetName, string cellAddress, string value)
        {
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Create new empty workbook.
            var workbook = new ExcelFile();

            // Add new sheet.
            var worksheet = workbook.Worksheets.Add(sheetName);

            // Write data to Excel cell.
            worksheet.Cells[cellAddress].Value = value;

            workbook.Save(excelFilePath);
        }
        
        public static void EditAnExistingExcelFile(string excelFilePath, string sheetName, string cellAddress, string value)
        {
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Load an Excel template.
            var workbook = ExcelFile.Load(excelFilePath);

            // Get template sheet.
            var worksheet = workbook.Worksheets[sheetName];

            // Write data to Excel cell.
            worksheet.Cells[cellAddress].Value = value;

            workbook.Save(excelFilePath);
        }
    }
}