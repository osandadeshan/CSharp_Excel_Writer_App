using System;
using System.Reflection;
using ExcelWriterApp.Util;

namespace ExcelWriterApp
{
    internal static class Program
    {
        private static void Main()
        {
            // Write to an Excel file without any styling
            BasicExcelWriter.Write();
            
            // Write to an Excel file without any styling
            StylesEmbeddedExcelWriter.Write();
            
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0, path.LastIndexOf("bin", 
                StringComparison.Ordinal))).LocalPath;
            var newExcelFilePath = new Uri(projectPath).LocalPath + @"\New_Simple_Excel_From_Util.xlsx";
            var editExcelFilePath = new Uri(projectPath).LocalPath + @"\Edit_Simple_Excel_From_Util.xlsx";
            
            // Write to an Excel file using the Util class
            ExcelWriter.WriteAsNewExcelFile(newExcelFilePath, "TestData", "A1", "Name");
            
            // Edit an existing Excel file using the Util class
            ExcelWriter.EditAnExistingExcelFile(newExcelFilePath, "TestData", "A1", "Name");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "B1", "Age");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "A2", "Osanda Nimalarathna");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "B2", "29");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "A3", "Eranga Nimalarathna");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "B3", "28");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "A4", "Nadeesha Nimalarathna");
            ExcelWriter.EditAnExistingExcelFile(editExcelFilePath, "TestData", "B4", "31");
        }
    }
}