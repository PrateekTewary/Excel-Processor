using System;
using System.Data;
using System.IO;
using ClosedXML.Excel;
class ExcelProcessor
{
    static void Main()
    {
        string inputFilePath = "Alumni Details.xlsx"; // Change to your actual input file path
        string outputFilePath = "processed_alumni.xlsx";

        if (!File.Exists(inputFilePath))
        {
            Console.WriteLine("Input file not found!");
            return;
        }

        using (var workbook = new XLWorkbook(inputFilePath))
        {
            var inputSheet = workbook.Worksheet(1);
            var outputWorkbook = new XLWorkbook();
            int lastRow = inputSheet.LastRowUsed().RowNumber();

            for (int row = 2; row <= lastRow; row++)  // Assuming row 1 is header
            {
                var rollNumber = inputSheet.Cell(row, 6).GetString();  // Assuming roll number is in column 1

                string yearSuffix = rollNumber.Split('/')[2]; // Extract year (last two digits)
                string year = "20" + yearSuffix; // Convert to full year (e.g., 19 -> 2019)

                if (!outputWorkbook.Worksheets.Contains(year))
                {
                    var newSheet = outputWorkbook.Worksheets.Add(year);
                    newSheet.Cell(1, 1).Value = "Timestamp"; // Add headers
                    newSheet.Cell(1, 2).Value = "Email Address"; // Assuming Name in column 2
                    newSheet.Cell(1, 3).Value = "Name";
                    newSheet.Cell(1, 4).Value = "Course";
                    newSheet.Cell(1, 5).Value = "Course Completion Year";
                    newSheet.Cell(1, 6).Value = "Institute Roll Number";
                    newSheet.Cell(1, 7).Value = "Contact Number";
                    newSheet.Cell(1, 8).Value = "Personal Email";
                    newSheet.Cell(1, 9).Value = "LinkedIn Profile URL";
                    newSheet.Cell(1, 10).Value = "Current Organization (Company/University)";
                    newSheet.Cell(1, 11).Value = "Current Position";
                }

                var outputSheet = outputWorkbook.Worksheet(year);
                int newRow = outputSheet.LastRowUsed()?.RowNumber() + 1 ?? 2;

                outputSheet.Cell(newRow, 1).Value = inputSheet.Cell(row, 1).Value;
                outputSheet.Cell(newRow, 2).Value = inputSheet.Cell(row, 2).Value;
                outputSheet.Cell(newRow, 3).Value = inputSheet.Cell(row, 3).Value;
                outputSheet.Cell(newRow, 4).Value = inputSheet.Cell(row, 4).Value;
                outputSheet.Cell(newRow, 5).Value = inputSheet.Cell(row, 5).Value;
                outputSheet.Cell(newRow, 6).Value = inputSheet.Cell(row, 6).Value;
                outputSheet.Cell(newRow, 7).Value = inputSheet.Cell(row, 7).Value;
                outputSheet.Cell(newRow, 8).Value = inputSheet.Cell(row, 8).Value;
                outputSheet.Cell(newRow, 9).Value = inputSheet.Cell(row, 9).Value;
                outputSheet.Cell(newRow, 10).Value =inputSheet.Cell(row, 10).Value;
                outputSheet.Cell(newRow, 11).Value =inputSheet.Cell(row, 11).Value;
            }

            outputWorkbook.SaveAs(outputFilePath);
            Console.WriteLine($"Processed data saved to {outputFilePath}");
        }
    }
}