﻿using System;
using System.Reflection;
using GemBox.Spreadsheet;

namespace ExcelWriterApp
{
    public class StylesEmbeddedExcelWriter
    {
        public void Write()
        {
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Create new empty workbook.
            var workbook = new ExcelFile();

            // Add new sheet.
            var worksheet = workbook.Worksheets.Add("Skyscrapers");

            // Write title to Excel cell.
            worksheet.Cells["A1"].Value = "List of tallest buildings (2021):";

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

            // // Set row formatting.
            worksheet.Rows["1"].Style = workbook.Styles[BuiltInCellStyleName.Heading1];

            // // Set columns width.
            worksheet.Columns["A"].SetWidth(8, LengthUnit.ZeroCharacterWidth); // Rank
            worksheet.Columns["B"].SetWidth(30, LengthUnit.ZeroCharacterWidth); // Building
            worksheet.Columns["C"].SetWidth(16, LengthUnit.ZeroCharacterWidth); // City
            worksheet.Columns["D"].SetWidth(20, LengthUnit.ZeroCharacterWidth); // Country
            worksheet.Columns["E"].SetWidth(9, LengthUnit.ZeroCharacterWidth); // Metric
            worksheet.Columns["F"].SetWidth(11, LengthUnit.ZeroCharacterWidth); // Imperial
            worksheet.Columns["G"].SetWidth(9, LengthUnit.ZeroCharacterWidth); // Floors
            worksheet.Columns["H"].SetWidth(9, LengthUnit.ZeroCharacterWidth); // Built (Year)
            worksheet.Columns["I"].SetWidth(4, LengthUnit.ZeroCharacterWidth); // Top 10
            worksheet.Columns["J"].SetWidth(5, LengthUnit.ZeroCharacterWidth); // Top 20

            // Write header data to Excel cells.
            for (int col = 0; col < skyscrapers.GetLength(1); col++)
                worksheet.Cells[3, col].Value = skyscrapers[0, col];
            worksheet.Cells["E3"].Value = "Height";

            worksheet.Cells.GetSubrange("A3:A4").Merged = true; // Rank
            worksheet.Cells.GetSubrange("B3:B4").Merged = true; // Building
            worksheet.Cells.GetSubrange("C3:C4").Merged = true; // City
            worksheet.Cells.GetSubrange("D3:D4").Merged = true; // Country
            worksheet.Cells.GetSubrange("E3:F3").Merged = true; // Height
            worksheet.Cells.GetSubrange("G3:G4").Merged = true; // Floors
            worksheet.Cells.GetSubrange("H3:H4").Merged = true; // Built (Year)

            // // Set header cells formatting.
            var style = new CellStyle();
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            style.VerticalAlignment = VerticalAlignmentStyle.Center;
            style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(237, 125, 49));
            style.Font.Weight = ExcelFont.BoldWeight;
            style.Font.Color = SpreadsheetColor.FromName(ColorName.White);
            style.WrapText = true;
            style.Borders.SetBorders(MultipleBorders.Right | MultipleBorders.Top,
                SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
            worksheet.Cells.GetSubrange("A3:H4").Style = style;

            // // Write "Top 10" cells.
            style = new CellStyle();
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            style.VerticalAlignment = VerticalAlignmentStyle.Center;
            style.Font.Weight = ExcelFont.BoldWeight;
            var mergedRange = worksheet.Cells.GetSubrange("I5:I14");
            mergedRange.Merged = true;
            mergedRange.Value = "T o p   1 0";
            style.Rotation = -90;
            style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(198, 239, 206));
            mergedRange.Style = style;

            // // Write "Top 20" cells.
            mergedRange = worksheet.Cells.GetSubrange("J5:J24");
            mergedRange.Merged = true;
            mergedRange.Value = "T o p   2 0";
            style.IsTextVertical = true;
            style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(255, 235, 156));
            mergedRange.Style = style;
            mergedRange = worksheet.Cells.GetSubrange("I15:I24");
            mergedRange.Merged = true;
            mergedRange.Style = style;

            // Write sample data and formatting to Excel cells.
            for (int row = 0; row < skyscrapers.GetLength(0) - 1; row++)
            {
                for (int col = 0; col < skyscrapers.GetLength(1); col++)
                {
                    var cell = worksheet.Cells[row + 4, col];
                    cell.Value = skyscrapers[row + 1, col];

                    cell.Style.Borders[IndividualBorder.Right].LineStyle = LineStyle.Thin;

                    if (row % 2 == 0)
                        cell.Style.FillPattern.SetSolid(SpreadsheetColor.FromArgb(221, 235, 247));

                    if (col == 0)
                        cell.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    if (col > 3)
                        cell.Style.Font.Name = "Courier New";
                    if (col == 4)
                        cell.Style.NumberFormat = "#\" m\"";
                    if (col == 5)
                        cell.Style.NumberFormat = "#\" ft\"";
                }
            }

            worksheet.Cells.GetSubrange("A5", "J24").Style.Borders.SetBorders(
                MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Medium);
            worksheet.Cells.GetSubrange("A3", "H4").Style.Borders.SetBorders(
                MultipleBorders.Vertical | MultipleBorders.Top, SpreadsheetColor.FromName(ColorName.Black),
                LineStyle.Medium);
            worksheet.Cells.GetSubrange("A5", "I14").Style.Borders.SetBorders(
                MultipleBorders.Bottom | MultipleBorders.Right, SpreadsheetColor.FromName(ColorName.Black),
                LineStyle.Medium);

            worksheet.PrintOptions.FitWorksheetWidthToPages = 1;

            // Save workbook as an Excel file.
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0, path.LastIndexOf("bin", 
                StringComparison.Ordinal))).LocalPath;
            var excelFilePath = new Uri(projectPath).LocalPath + @"\Excel_With_Styles.xlsx";
            
            workbook.Save(excelFilePath);
        }
    }
}