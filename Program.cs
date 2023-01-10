using System;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using static System.Runtime.InteropServices.JavaScript.JSType;

//this is a code junkyard now

internal class Program
{

    public static void Main(string[] args)
    {
        //this code makes the program usable for any file by taking the file path as a command line argument
        Console.WriteLine("Enter file by dragging Excel you wish to filter into this terminal window.");
        string fileName = Console.ReadLine();
        //string[] = Console.ReadLine();
        //assign excel doc we want to read to a var
        using (var package = new ExcelPackage(new FileInfo(fileName:fileName))) //
        {
            /////DEFINE FIRST SHEET
            var firstSheet = package.Workbook.Worksheets["Sheet1"]; //selects sheet

            Console.WriteLine("\nWelcome to Services Tracker, this application is made to filter large GCSSA spreadsheets down to readable size.");
            Console.WriteLine("You'll be prompted to enter your company, please enter as an abbreviation for the program to function correctly.");
            Console.WriteLine("Once you've entered your company in uppercase, you'll enter the type of date you want to be shown.");
            Console.WriteLine("The program will run for 3-4 minutes and create a second sheet on the raw data GCSSA spreadsheet you downloaded");
            Console.WriteLine("When this terminal closes, you'll be able to open the file and see the filtered second sheet.");
            Console.WriteLine("When complete, rename the resulting file with two sheets, so 'services.xlxs' can be used for new files. Enjoy!\n");

            ////GET USER INPUTS////////
            Console.WriteLine("Enter your Company (Ex: A,B,C,HHC,FSC)\n");
            var userInput = Console.ReadLine();
            Console.WriteLine("\n");
            Console.WriteLine("What type of date would you like (Ex: early, planned, or late)\n");
            var userDate = Console.ReadLine();

            //////DEFINE SHEET 2
            var outputSheet = package.Workbook.Worksheets.Add("Sheet2");


            /////STYLE SHEET 2
            outputSheet.TabColor = System.Drawing.Color.Black;
            outputSheet.DefaultRowHeight = 12;
            outputSheet.Row(1).Height = 20;
            outputSheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            outputSheet.Row(1).Style.Font.Bold = true;
            outputSheet.Row(1).Style.Border.BorderAround(ExcelBorderStyle.Hair);
            outputSheet.Column(1).Width = 20;
            outputSheet.Column(2).Width = 20;
            outputSheet.Column(3).Width = 20;
            outputSheet.Column(4).Width = 40;
            outputSheet.Column(5).Width = 20;
            outputSheet.Column(6).Width = 20;
            outputSheet.Column(7).Width = 20;

            //////ADD HEADERS
            outputSheet.Cells["A1"].Value = "UIC";
            outputSheet.Cells["B1"].Value = "Admin Number";
            outputSheet.Cells["C1"].Value = "Model Number";
            outputSheet.Cells["D1"].Value = "Description";
            outputSheet.Cells["G1"].Value = "Overdue or Open";

            switch (userDate)
            {
                case "early":
                    outputSheet.Cells["E1"].Value = "Early Date";
                    outputSheet.Cells["F1"].Value = "Days Until Early Date";
                    break;

                case "planned":
                    outputSheet.Cells["E1"].Value = "Planned Date";
                    outputSheet.Cells["F1"].Value = "Days Until Planned Date";
                    break;

                case "late":
                    outputSheet.Cells["E1"].Value = "Late Date";
                    outputSheet.Cells["F1"].Value = "Days Until Late Date";
                    break;
            }

            ///below code makes program compatible with all GCSS layouts

            var uicColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Main work center")
                .Start
                .Column;
            var adminColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Admin No.")
                .Start
                .Column;
            var modelColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Model number")
                .Start
                .Column;
            var descColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Description of technical object")
                .Start
                .Column;
            var earlyColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Early Date")
                .Start
                .Column;
            var plannedColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "PlanDate MaintCall")
                .Start
                .Column;
            var lateColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Late Date")
                .Start
                .Column;
            var overdueColumn =
                firstSheet
                .Cells["A1:Z1"]
                .First(c => c.Value.ToString() == "Completion Status")
                .Start
                .Column;

            var dayToday = DateTime.Now.ToString("MM/dd/yyyy");

            switch (userDate)
            {
                case "early" or "EARLY" or "Early":
                    var start = firstSheet.Dimension.Start;
                    var end = firstSheet.Dimension.End;
                    int dateRow = earlyColumn;

                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        var uicString = firstSheet.Cells[row, 1].Value.ToString();
                        char uic = uicString[4];
                            switch (userInput)
                            {
                                case "A":
                                    if (uic is 'A')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open

                                }
                                    package.Save();
                                    break;
                                case "B":
                                    if (uic is 'B')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                }
                                    package.Save();
                                    break;
                            case "C":
                                if (uic is 'C')
                                {
                                    firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                    firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                    firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                    firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                    firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                                    firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                }
                                package.Save();
                                break;
                            case "HHC":
                                if (uic is 'T')
                                {
                                    firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                    firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                    firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                    firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                    firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                                    firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                }
                                package.Save();
                                break;
                            case "FSC":
                                if (uic is 'D' or 'E' or 'F' or 'G' or 'H' or 'I' or 'J' or 'K' or 'L')
                                {
                                    firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                    firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                    firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                    firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                    firstSheet.Cells[row, dateRow].Copy(outputSheet.Cells[row, 5]); //type of Date
                                    firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                }
                                package.Save();
                                break;
                        }
                    }
                    break;
                case "late":
                    {
                        var startLate = firstSheet.Dimension.Start;
                        var endLate = firstSheet.Dimension.End;
                        int dateRowLate = lateColumn;

                        for (int row = startLate.Row; row <= endLate.Row; row++)
                        {
                            var uicString = firstSheet.Cells[row, 1].Value.ToString();
                            char uic = uicString[4];
                            switch (userInput)
                            {
                                case "A":
                                    if (uic is 'A')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowLate].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open

                                    }
                                    package.Save();
                                    break;
                                case "B":
                                    if (uic is 'B')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowLate].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "C":
                                    if (uic is 'C')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowLate].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "HHC":
                                    if (uic is 'T')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowLate].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "FSC":
                                    if (uic is 'D' or 'E' or 'F' or 'G' or 'H' or 'I' or 'J' or 'K' or 'L')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowLate].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                            }
                        }
                    }
                    break;
                case "planned":
                    {
                        var startPlan = firstSheet.Dimension.Start;
                        var endPlan = firstSheet.Dimension.End;
                        int dateRowPlan = plannedColumn;

                        for (int row = startPlan.Row; row <= endPlan.Row; row++)
                        {
                            var uicString = firstSheet.Cells[row, 1].Value.ToString();
                            char uic = uicString[4];
                            switch (userInput)
                            {
                                case "A":
                                    if (uic is 'A')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowPlan].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "B":
                                    if (uic is 'B')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowPlan].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "C":
                                    if (uic is 'C')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowPlan].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "HHC":
                                    if (uic is 'T')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowPlan].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                                case "FSC":
                                    if (uic is 'D' or 'E' or 'F' or 'G' or 'H' or 'I' or 'J' or 'K' or 'L')
                                    {
                                        firstSheet.Cells[row, uicColumn].Copy(outputSheet.Cells[row, 1]); //UIC
                                        firstSheet.Cells[row, adminColumn].Copy(outputSheet.Cells[row, 2]); //adminNum
                                        firstSheet.Cells[row, modelColumn].Copy(outputSheet.Cells[row, 3]); //modelNum
                                        firstSheet.Cells[row, descColumn].Copy(outputSheet.Cells[row, 4]); //Description
                                        firstSheet.Cells[row, dateRowPlan].Copy(outputSheet.Cells[row, 5]); //type of Date
                                        firstSheet.Cells[row, overdueColumn].Copy(outputSheet.Cells[row, 7]); //overdue or open
                                    }
                                    package.Save();
                                    break;
                            }
                        }
                    }
                    break;
            }
            outputSheet.Cells["A2:G4000"].Sort(4, false);

            var startCol6 = outputSheet.Row(2);
            var endCol6 = outputSheet.Dimension.End;
            for (int row = startCol6.Row; row <= endCol6.Row; row++)
            {
                if (outputSheet.Cells[row, 5].Value is not null)
                 {
                     var cellDate = outputSheet.Cells[row, 5].Value.ToString();
                     DateTime startTime = DateTime.Parse(cellDate);
                     DateTime dt = DateTime.Parse(dayToday);
                     TimeSpan t = startTime - dt;
                     outputSheet.Cells[row, 6].Value = (int)t.TotalDays;
                 }
            }
            package.Save();
          
        } }
    

}
