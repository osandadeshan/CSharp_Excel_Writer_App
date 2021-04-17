using System;
using System.Reflection;
using GemBox.Spreadsheet;

namespace ExcelWriterApp
{
    public static class BasicExcelWriter
    {
        public static void Write()
        {
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Create new empty workbook.
            var workbook = new ExcelFile();

            // Add new sheet.
            var worksheet = workbook.Worksheets.Add("Skyscrapers");

            // Write title to Excel cell.
            // worksheet.Cells["A1"].Value = "List of tallest buildings (2021):";

            // Tabular sample data for writing into an Excel file.
            var skyscrapers = new object[,]
            {
                {"Rank", "Building", "City", "Country", "Metric", "Imperial", "Floors", "Built (Year)"},
                {1, "Burj Khalifa", "Dubai", "United Arab Emirates", 828, 2717, 163, 2010},
                {2, "Shanghai Tower", "Shanghai", "China", 632, 2073, 128, 2015},
                {3, "Abraj Al-Bait Clock Tower", "Mecca", "Saudi Arabia", 601, 1971, 120, 2012},
                {4, "Ping An Finance Centre", "Shenzhen", "China", 599, 1965, 115, 2017},
                {5, "Lotte World Tower", "Seoul", "South Korea", 554.5, 1819, 123, 2016},
                {6, "One World Trade Center", "New York City", "United States", 541.3, 1776, 104, 2014},
                {7, "Guangzhou CTF Finance Centre", "Guangzhou", "China", 530, 1739, 111, 2016},
                {8, "Tianjin CTF Finance Centre", "Tianjin", "China", 530, 1739, 98, 2019},
                {9, "China Zun", "Beijing", "China", 528, 1732, 108, 2018},
                {10, "Taipei 101", "Taipei", "Taiwan", 508, 1667, 101, 2004},
                {11, "Shanghai World Financial Center", "Shanghai", "China", 492, 1614, 101, 2008},
                {12, "International Commerce Centre", "Hong Kong", "China", 484, 1588, 118, 2010},
                {13, "Central Park Tower", "New York City", "United States", 472, 1550, 98, 2020},
                {14, "Lakhta Center", "St. Petersburg", "Russia", 462, 1516, 86, 2019},
                {15, "Landmark 81", "Ho Chi Minh City", "Vietnam", 461.2, 1513, 81, 2018},
                {16, "Changsha IFS Tower T1", "Changsha", "China", 452.1, 1483, 88, 2018},
                {17, "Petronas Tower 1", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998},
                {18, "Petronas Tower 2", "Kuala Lumpur", "Malaysia", 451.9, 1483, 88, 1998},
                {19, "Zifeng Tower", "Nanjing", "China", 450, 1476, 89, 2010},
                {20, "Suzhou IFS", "Suzhou", "China", 450, 1476, 98, 2019}
            };

            // Write header data to Excel cells.
            for (var col = 0; col < skyscrapers.GetLength(1); col++)
                worksheet.Cells[0, col].Value = skyscrapers[0, col];

            // Write sample data and formatting to Excel cells.
            for (var row = 0; row < skyscrapers.GetLength(0); row++)
            {
                for (var col = 0; col < skyscrapers.GetLength(1); col++)
                {
                    var cell = worksheet.Cells[row, col];
                    cell.Value = skyscrapers[row, col];
                }
            }

            worksheet.PrintOptions.FitWorksheetWidthToPages = 1;

            // Save workbook as an Excel file.
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0, path.LastIndexOf("bin", 
                StringComparison.Ordinal))).LocalPath;
            var excelFilePath = new Uri(projectPath).LocalPath + @"\Basic_Excel.xlsx";
            
            workbook.Save(excelFilePath);
        }
    }
}