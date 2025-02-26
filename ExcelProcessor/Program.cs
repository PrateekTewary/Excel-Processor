using System;
using System.Data;
using System.IO;
using System.Linq;
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
            var tempWorkBook = new XLWorkbook();
            int lastRow = inputSheet.LastRowUsed().RowNumber();
            Dictionary<string, 
                List<(string Timestamp,
                 string EmailAddress,
                 string Name, 
                 string Course, 
                 string CourseCompletionYear, 
                 string InstituteRollNumber, 
                 string ContactNumber, 
                 string PersonalEmail, 
                 string LinkedInProfileURL, 
                 string CurrentOrganization, 
                 string CurrentPosition)>> sheetData = new();

            for (int row = 2; row <= lastRow; row++)  // Assuming row 1 is header
            {
                var rollNumber = inputSheet.Cell(row, 6).GetString();  // Assuming roll number is in column 1

                string yearSuffix = rollNumber.Split('/')[2]; // Extract year (last two digits)
                string year = "20" + yearSuffix; // Convert to full year (e.g., 19 -> 2019)
                
                if (!sheetData.ContainsKey(year))
                    sheetData[year] = new List<(string Timestamp,
                                string EmailAddress,
                                string Name, 
                                string Course, 
                                string CourseCompletionYear, 
                                string InstituteRollNumber, 
                                string ContactNumber, 
                                string PersonalEmail, 
                                string LinkedInProfileURL, 
                                string CurrentOrganization, 
                                string CurrentPosition)>();
                
                sheetData[year].Add((
                    inputSheet.Cell(row, 1).GetValue<string>(),
                    inputSheet.Cell(row, 2).GetValue<string>(),
                    inputSheet.Cell(row, 3).GetValue<string>(),
                    inputSheet.Cell(row, 4).GetValue<string>(),
                    inputSheet.Cell(row, 5).GetValue<string>(),
                    inputSheet.Cell(row, 6).GetValue<string>(),
                    inputSheet.Cell(row, 7).GetValue<string>(),
                    inputSheet.Cell(row, 8).GetValue<string>(),
                    inputSheet.Cell(row, 9).GetValue<string>(),
                    inputSheet.Cell(row, 10).GetValue<string>(),
                    inputSheet.Cell(row, 11).GetValue<string>()
                ));
            }

            foreach (var year in sheetData.Keys.OrderBy(y => int.Parse(y)))
            {
                // Sort rows by middle value numerically
                var sortedData = sheetData[year]
                    .Where(entry => !string.IsNullOrWhiteSpace(entry.InstituteRollNumber)) // Ensure roll number is not empty
                    .OrderBy(entry => int.Parse(entry.InstituteRollNumber.Split('/')[1])) // Sorting by middle value
                    .ToList();
                if (!tempWorkBook.Worksheets.Contains(year))
                {
                    tempWorkBook.Worksheets.Add(year);
                }
                var outputSheet = tempWorkBook.Worksheet(year);
                SetSheetHeaders(outputSheet);
                int newRow = 2;
                foreach (var record in sortedData)
                {
                    PopulateSheetRow(outputSheet,newRow,record);
                    ++newRow;
                }
            }

            // Sort sheets by name (year) and create a new final workbook
            
             using (var finalWorkbook = new XLWorkbook())
            {
                foreach (var year in sheetData.Keys.OrderBy(y => int.Parse(y)))
                {
                    tempWorkBook.Worksheet(year).CopyTo(finalWorkbook, year);
                }

                finalWorkbook.SaveAs(outputFilePath);
            }

            Console.WriteLine($"Processed data saved to {outputFilePath} with sorted sheets.");
        }
    }

    static void PopulateSheetRow(IXLWorksheet sheet, int row,
                                    (string Timestamp, 
                                     string EmailAddress,
                                     string Name, 
                                     string Course, 
                                     string CourseCompletionYear, 
                                     string InstituteRollNumber, 
                                     string ContactNumber, 
                                     string PersonalEmail, 
                                     string LinkedInProfileURL, 
                                     string CurrentOrganization, 
                                     string CurrentPosition) alumni)
    {
        sheet.Cell(row, 1).Value =  alumni.Timestamp;
        sheet.Cell(row, 2).Value =  alumni.EmailAddress;
        sheet.Cell(row, 3).Value =  alumni.Name;
        sheet.Cell(row, 4).Value =  alumni.Course;
        sheet.Cell(row, 5).Value =  alumni.CourseCompletionYear;
        sheet.Cell(row, 6).Value =  alumni.InstituteRollNumber;
        sheet.Cell(row, 7).Value =  alumni.ContactNumber;
        sheet.Cell(row, 8).Value =  alumni.PersonalEmail;
        sheet.Cell(row, 9).Value =  alumni.LinkedInProfileURL;
        sheet.Cell(row, 10).Value = alumni.CurrentOrganization;
        sheet.Cell(row, 11).Value = alumni.CurrentPosition;
        sheet.Columns().AdjustToContents();
    }
    static void SetSheetHeaders(IXLWorksheet sheet)
    {
        sheet.Cell(1, 1).Value = "Timestamp"; // Add headers
        sheet.Cell(1, 1).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;

        sheet.Cell(1, 2).Value = "Email Address"; // Assuming Name in column 2
        sheet.Cell(1, 2).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 3).Value = "Name";
        sheet.Cell(1, 3).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 4).Value = "Course";
        sheet.Cell(1, 4).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 5).Value = "Course Completion Year";
        sheet.Cell(1, 5).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 6).Value = "Institute Roll Number";
        sheet.Cell(1, 6).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 7).Value = "Contact Number";
        sheet.Cell(1, 7).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 8).Value = "Personal Email";
        sheet.Cell(1, 8).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 9).Value = "LinkedIn Profile URL";
        sheet.Cell(1, 9).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 10).Value = "Current Organization (Company/University)";
        sheet.Cell(1, 10).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
        
        sheet.Cell(1, 11).Value = "Current Position";
        sheet.Cell(1, 11).Style.Font.SetFontSize(15).Font.SetFontName("Roboto").Fill.SetBackgroundColor(XLColor.Blue).Font.Bold = true;
    }
}