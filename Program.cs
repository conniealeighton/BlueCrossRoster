using System;
using OfficeOpenXml;
using System.IO;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using System.Drawing.Imaging;
using System.Reflection;
using OfficeOpenXml.Sorting;
using System.Security.Cryptography.X509Certificates;
using System.Globalization;
using System.ComponentModel.DataAnnotations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace ExcelDataTransfer
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set to NonCommercial if appropriate

            string spreadsheetAPath = "../Downloads/BlueCrossRosterFile6WeePreProd.xlsx";

            // Load Spreadsheet A
            using (var packageA = new ExcelPackage(new FileInfo(spreadsheetAPath)))
            {
                //Get all spreadsheets by name
                var BlueCrossRoster = packageA.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "BlueCrossRoster");
                var LookoutRoster = packageA.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "LookoutRoster");
                var workersSheet = packageA.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "Workers");
                int lookoutRosterCount = 4;


                int rowCountA = BlueCrossRoster.Dimension.Rows;


                // Iterate through each row in Spreadsheet A
                for (int rowA = 2; rowA <= rowCountA; rowA++) // Assuming header in the first row
                {
                    //Get Membership type
                    int funderColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Funding Body");
                    string funder = GetFunder(BlueCrossRoster.Cells[rowA, funderColumnID].Value?.ToString());

                    if (funder == "unknown")
                    {
                        Console.WriteLine("unknown funder");
                        continue;
                    }

                    // Get Member Name from Blue Cross Roster
                    int nameColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Client Name");
                    string memberName = GetReorderedName(BlueCrossRoster.Cells[rowA, nameColumnID].Value?.ToString());

                    //get helper info from Blue Cross Roster
                    int blueCrossWorkerNameColumn = GetColumnNumberByHeaderName(BlueCrossRoster, "Carer Name");
                    string workerName = GetReorderedName(BlueCrossRoster.Cells[rowA, blueCrossWorkerNameColumn].Value?.ToString());

                    // Get helper name from implimentation workbook
                    int lookoutWorkerNameColumnID = GetColumnNumberByHeaderName(workersSheet, "full_name");
                    int lookoutHelperID = GetColumnNumberByHeaderName(workersSheet, "helper_id");
                    string workerID;

                    //Get rates and membership sheets for member's funder
                    var membershipSheet = packageA.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == ("3 - Memberships - " + funder)); ;
                    var ratesSheet = packageA.Workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == funder);

                    int membershipColumnID = GetColumnNumberByHeaderName(membershipSheet, "membership_id");
                    int profileColumnID = GetColumnNumberByHeaderName(membershipSheet, "client_profile_id");
                    int membershipNameColumnID = GetColumnNumberByHeaderName(membershipSheet, "client_full_name");
                    string membershipID;
                    string profileID;

                    //Get week commensing, ending and weekly intervals 
                    int weekCommensingColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Start Date");
                    DateTime weekCommensing = DateTime.Parse(BlueCrossRoster.Cells[rowA, weekCommensingColumnID].Value?.ToString());

                    //Get end date
                    string endDate = "";
                    int descColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Desc");
                    string desc = BlueCrossRoster.Cells[rowA, descColumnID].Value?.ToString();

                    //If the description specifies "no end date" then we know there never is one
                    if (desc.Contains("no end date."))
                    {
                        endDate = "Never";
                    }
                    //when there is an end date it can be found in the end date column of the blue cross roster
                    else
                    {
                        int endDateColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "End Date");
                        endDate = BlueCrossRoster.Cells[rowA, endDateColumnID].Value?.ToString();
                    }

                    //get frequency of visits
                    string pattern = @"Every(?: (\d{1,2})(?:st|nd|rd|th)?)? (\w+)";
                    int frequency = -1;
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(desc);
                    if (match.Success)
                    {
                        if (string.IsNullOrEmpty(match.Groups[1].Value))
                        {
                            // If there is no numeric frequency specified, default to 1
                            frequency = 1;
                        }

                        else if (int.TryParse(match.Groups[1].Value, out frequency))
                        {
                            // Successfully parsed the numeric frequency, returning it as an integer
                        }
                    }
                    else
                    {
                        Console.WriteLine("cannot match regex: " + desc);
                        continue;
                    }

                    //Get start date
                    int startDateColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Start Date");
                    DateTime blueCrossStartDate = DateTime.Parse(BlueCrossRoster.Cells[rowA, startDateColumnID].Value?.ToString());
                    DateTime startDate = getRuleStartDate(blueCrossStartDate, frequency);

                    // Find matching worker in worker sheet
                    int workerRow = Enumerable.Range(2, rowCountA - 1)
                        .FirstOrDefault(i => string.Equals(workerName, workersSheet.Cells[i, lookoutWorkerNameColumnID].Value?.ToString(),
                        StringComparison.OrdinalIgnoreCase));

                    string lookoutWorkerName = "";

                    if (workerName.Contains("PERM VACANT") || workerName.Contains("LEAVE VACANT"))
                    {
                        workerID = "-1";
                    }
                    else if (workerRow != 0)
                    {
                        workerID = workersSheet.Cells[workerRow, lookoutHelperID].Value?.ToString();
                    }
                    else
                    {
                        Console.WriteLine("cannot match worker. Name: " + workerName);
                        continue;
                    }
                    if (workerID != "-1")
                    {
                        lookoutWorkerName = workersSheet.Cells[workerRow, lookoutWorkerNameColumnID].Value?.ToString();
                    }
                    else lookoutWorkerName = "Vacant";


                    // Find matching row in Membership Spreadsheet
                    int membershipRow = Enumerable.Range(2, rowCountA - 1)
                        .FirstOrDefault(i => string.Equals(memberName, membershipSheet.Cells[i, membershipNameColumnID].Value?.ToString(),
                        StringComparison.OrdinalIgnoreCase));

                    if (membershipRow != 0)
                    {
                        membershipID = membershipSheet.Cells[membershipRow, membershipColumnID].Value?.ToString();
                        profileID = membershipSheet.Cells[membershipRow, profileColumnID].Value?.ToString();
                    }

                    else
                    {
                        Console.WriteLine("cannot match member name. Funding scheme: " + funder + ". +Member Name: " + memberName);
                        continue;
                    }

                    //Get day(s) of rosted rule
                    int rosterDaysColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Days");
                    string rosteredDays = BlueCrossRoster.Cells[rowA, rosterDaysColumnID].Value?.ToString();
                    string[] daysArray = rosteredDays.Split(',');
                    string newDays = "";
                    bool hasSaturday = false;
                    bool hasSunday = false;

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        if (daysArray[i] != null && daysArray[i].Equals("Saturday", StringComparison.OrdinalIgnoreCase))
                        {
                            // If the value is found, remove it from the original array
                            string removedValue = daysArray[i];
                            daysArray[i] = null;

                            hasSaturday = true;
                        }
                    }
                    
                    

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        if (daysArray[i] != null && daysArray != null && daysArray[i].Equals("Sunday", StringComparison.OrdinalIgnoreCase))
                        {
                            // If the value is found, remove it from the original array
                            string removedValue = daysArray[i];
                            daysArray[i] = null;

                            hasSunday = true;
                        }
                    }
                    

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        if((i + 1) == daysArray.Length) newDays += daysArray[i];
                        else newDays += daysArray[i] + ", ";
                        rosteredDays = newDays;
                    }

                    if(daysArray.Length == 0)
                    {
                        rosteredDays = "";
                    }


                    //Get start and end time 
                    int startTimeColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Start Time");
                    var startTime = BlueCrossRoster.Cells[rowA, startTimeColumnID].Text;
                    TimeSpan startTimeParse = TimeSpan.ParseExact(startTime, "hh\\:mm", System.Globalization.CultureInfo.InvariantCulture);


                    int endTimeColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "End Time");
                    string endTime = BlueCrossRoster.Cells[rowA, endTimeColumnID].Text;
                    TimeSpan endTimeParse = TimeSpan.ParseExact(endTime, "hh\\:mm", System.Globalization.CultureInfo.InvariantCulture);

                    //Get Service
                    int serviceColumnID = GetColumnNumberByHeaderName(BlueCrossRoster, "Service");
                    string blueCrossService = BlueCrossRoster.Cells[rowA, serviceColumnID].Value?.ToString();

                    int lookoutServiceColumnID = GetColumnNumberByHeaderName(ratesSheet, "Lookout");
                    int rateColumnID = GetColumnNumberByHeaderName(ratesSheet, "Rate");
                    int blueCrossNameColumnID = GetColumnNumberByHeaderName(ratesSheet, "BlueCross");
                    string rate = "";
                    int ratesRow = Enumerable.Range(2, rowCountA - 1)
                    .FirstOrDefault(i => string.Equals(blueCrossService, ratesSheet.Cells[i, blueCrossNameColumnID].Value?.ToString(),
                    StringComparison.OrdinalIgnoreCase));
                    string lookoutService = ratesSheet.Cells[ratesRow, lookoutServiceColumnID].Value?.ToString();
                    string lookoutRate = ratesSheet.Cells[ratesRow, rateColumnID].Value?.ToString();

                    rate = getRate(funder, membershipSheet, membershipRow, memberName, blueCrossService,
                    lookoutService, ratesRow, rosteredDays, startTimeParse);


                    if (rate == "") continue;

                    if(daysArray.Length > 0 && rosteredDays != "") {

                    LookoutRoster.Cells[lookoutRosterCount, 3].Value = membershipID;
                    LookoutRoster.Cells[lookoutRosterCount, 4].Value = profileID;
                    LookoutRoster.Cells[lookoutRosterCount, 5].Value = memberName;
                    LookoutRoster.Cells[lookoutRosterCount, 6].Value = workerID;
                    LookoutRoster.Cells[lookoutRosterCount, 7].Value = lookoutWorkerName;
                    LookoutRoster.Cells[lookoutRosterCount, 11].Value = "FALSE";
                    LookoutRoster.Cells[lookoutRosterCount, 12].Value = startDate;
                    LookoutRoster.Cells[lookoutRosterCount, 13].Value = endDate;
                    LookoutRoster.Cells[lookoutRosterCount, 14].Value = rosteredDays;
                    LookoutRoster.Cells[lookoutRosterCount, 15].Value = frequency;
                    LookoutRoster.Cells[lookoutRosterCount, 16].Value = startTime;
                    LookoutRoster.Cells[lookoutRosterCount, 17].Value = endTime;
                    LookoutRoster.Cells[lookoutRosterCount, 18].Value = lookoutService;
                    LookoutRoster.Cells[lookoutRosterCount, 19].Value = rate;

                    if (funder == "HCP" && (lookoutService == "Companion Care" || lookoutService == "Domestic Assistance (Home Care)" ||
                    lookoutService == "Personal Care" || lookoutService == "Shopping - escorted" || lookoutService == "Transport"))
                    {
                        LookoutRoster.Cells[lookoutRosterCount, 8].Value = "Yes";
                        LookoutRoster.Cells[lookoutRosterCount, 9].Value = "HCP Travel per KM";
                    }
                    else
                    {
                        LookoutRoster.Cells[lookoutRosterCount, 8].Value = "No";
                    }

                    lookoutRosterCount++;
                    }

                    //Do Saturday
                    if(hasSaturday)
                    {
                        rate = getRate(funder, membershipSheet, membershipRow, memberName, blueCrossService,
                    lookoutService, ratesRow, "Saturday", startTimeParse);
                    

                    if (rate == "") continue;

                    LookoutRoster.Cells[lookoutRosterCount, 3].Value = membershipID;
                    LookoutRoster.Cells[lookoutRosterCount, 4].Value = profileID;
                    LookoutRoster.Cells[lookoutRosterCount, 5].Value = memberName;
                    LookoutRoster.Cells[lookoutRosterCount, 6].Value = workerID;
                    LookoutRoster.Cells[lookoutRosterCount, 7].Value = lookoutWorkerName;
                    LookoutRoster.Cells[lookoutRosterCount, 11].Value = "FALSE";
                    LookoutRoster.Cells[lookoutRosterCount, 12].Value = startDate;
                    LookoutRoster.Cells[lookoutRosterCount, 13].Value = endDate;
                    LookoutRoster.Cells[lookoutRosterCount, 14].Value = "Saturday";
                    LookoutRoster.Cells[lookoutRosterCount, 15].Value = frequency;
                    LookoutRoster.Cells[lookoutRosterCount, 16].Value = startTime;
                    LookoutRoster.Cells[lookoutRosterCount, 17].Value = endTime;
                    LookoutRoster.Cells[lookoutRosterCount, 18].Value = lookoutService;
                    LookoutRoster.Cells[lookoutRosterCount, 19].Value = rate;

                    if (funder == "HCP" && (lookoutService == "Companion Care" || lookoutService == "Domestic Assistance (Home Care)" ||
                    lookoutService == "Personal Care" || lookoutService == "Shopping - escorted" || lookoutService == "Transport"))
                    {
                        LookoutRoster.Cells[lookoutRosterCount, 8].Value = "Yes";
                        LookoutRoster.Cells[lookoutRosterCount, 9].Value = "HCP Travel per KM";
                    }
                    else
                    {
                        LookoutRoster.Cells[lookoutRosterCount, 8].Value = "No";
                    }

                    lookoutRosterCount++;
                    }

                    //Do Sunday
                    if(hasSunday)
                    {
                        rate = getRate(funder, membershipSheet, membershipRow, memberName, blueCrossService,
                    lookoutService, ratesRow, "Sunday", startTimeParse);
                    

                    if (rate == "") continue;

                    LookoutRoster.Cells[lookoutRosterCount, 3].Value = membershipID;
                    LookoutRoster.Cells[lookoutRosterCount, 4].Value = profileID;
                    LookoutRoster.Cells[lookoutRosterCount, 5].Value = memberName;
                    LookoutRoster.Cells[lookoutRosterCount, 6].Value = workerID;
                    LookoutRoster.Cells[lookoutRosterCount, 7].Value = lookoutWorkerName;
                    LookoutRoster.Cells[lookoutRosterCount, 11].Value = "FALSE";
                    LookoutRoster.Cells[lookoutRosterCount, 12].Value = startDate;
                    LookoutRoster.Cells[lookoutRosterCount, 13].Value = endDate;
                    LookoutRoster.Cells[lookoutRosterCount, 14].Value = "Sunday";
                    LookoutRoster.Cells[lookoutRosterCount, 15].Value = frequency;
                    LookoutRoster.Cells[lookoutRosterCount, 16].Value = startTime;
                    LookoutRoster.Cells[lookoutRosterCount, 17].Value = endTime;
                    LookoutRoster.Cells[lookoutRosterCount, 18].Value = lookoutService;
                    LookoutRoster.Cells[lookoutRosterCount, 19].Value = rate;

                    if (funder == "HCP" && (lookoutService == "Companion Care" || lookoutService == "Domestic Assistance (Home Care)" ||
                    lookoutService == "Personal Care" || lookoutService == "Shopping - escorted" || lookoutService == "Transport"))
                    {
                        LookoutRoster.Cells[lookoutRosterCount, 8].Value = "Yes";
                        LookoutRoster.Cells[lookoutRosterCount, 9].Value = "HCP Travel per KM";
                    }
                    else
                    {
                        LookoutRoster.Cells[lookoutRosterCount, 8].Value = "No";
                    }

                    lookoutRosterCount++;
                    }

                }

                packageA.Save();
            }
            Console.WriteLine("Data copied successfully.");

            static string getRate(string funder, ExcelWorksheet membershipSheet, int membershipRow, string memberName, string blueCrossService,
            string lookoutService, int ratesRow, string rosteredDays, TimeSpan startTimeParse)
            {
                if (funder == "CHSP")
                {
                    int cocontributionColumnID = GetColumnNumberByHeaderName(membershipSheet, "Contribution");
                    string cocontribution = membershipSheet.Cells[membershipRow, cocontributionColumnID].Value?.ToString();
                    if (cocontribution == "" || cocontribution == null)
                    {
                        Console.WriteLine("CHSP No contribution record for: " + memberName);
                        return "";
                    }

                    if (cocontribution == "part")
                    {
                        if (blueCrossService == "Personal Care")
                        {
                            return "CHSP - Personal Care (Co-Contribution-Part Pensioner or Self-Funded)";
                        }
                        else if (blueCrossService == "Home Care")
                        {
                            return "CHSP - Homecare (Co-Contribution-Part Pensioner or Self-Funded)";
                        }
                        else if (blueCrossService == "Respite")
                        {
                            return "CHSP - Respite (Co-Contribution-Part Pensioner or Self-Funded)";
                        }
                    }
                    else if (cocontribution == "full")
                    {
                        if (blueCrossService == "Personal Care")
                        {
                            return "CHSP - Personal Care (Co-Contribution-Full Pensioner)";
                        }
                        else if (blueCrossService == "Home Care")
                        {
                            return "CHSP - Homecare (Co-Contribution-Full Pensioner)";
                        }
                        else if (blueCrossService == "Respite")
                        {
                            return "CHSP - Respite (Co-Contribution-Full Pensioner)";
                        }
                    }
                    else
                    {
                        Console.WriteLine("Couldn't match service: " + blueCrossService);
                    }
                }

                if (funder == "HCP")
                {
                    if (ratesRow == 0)
                    {
                        Console.WriteLine("cannot match rates");
                        return "";
                    }

                    if (lookoutService == "Registered Nurse")
                    {
                        switch (rosteredDays)
                        {
                            case "Saturday":
                                return "Registered Nurse - Saturday";
                                break;
                            case "Sunday":
                                return "Registered Nurse - Sunday";
                                break;
                            default:
                                return "Registered Nurse - Weekday";
                                break;
                        }
                    }
                    else if (lookoutService == "Personal Care" || lookoutService == "Domestic Assistance (Home Care)" ||
                    lookoutService == "Respite" || lookoutService == "Transport")
                    {
                        switch (rosteredDays)
                        {
                            case "Saturday" when startTimeParse > new TimeSpan(20, 0, 0): // 20:00 or 8pm
                                return "Personal Care - Night Shift - Saturday";
                                break;

                            case "Sunday" when startTimeParse > new TimeSpan(20, 0, 0): // 20:00 or 8pm
                                return "Personal Care - Night Shift - Sunday";
                                break;

                            case "Saturday":
                                return "Personal Care - Saturday";
                                break;

                            case "Sunday":
                                return "Personal Care - Sunday";
                                break;

                            case var _ when startTimeParse > new TimeSpan(20, 0, 0): // 20:00 or 8pm
                                return "Personal Care - Night Shift - Weekday";
                                break;

                            default:
                                return "Personal Care / Home Care - Weekday";
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("HCP rate Unassigned");
                        return "";
                    }
                }
                if (funder == "Private")
                {
                    if (ratesRow == 0)
                    {
                        Console.WriteLine("cannot match rates");
                        return "";
                    }

                    if (lookoutService == "Personal Care" || lookoutService == "Domestic Assistance (Home Care)" ||
                    lookoutService == "Companion Care" || lookoutService == "Transport")
                    {
                        switch (rosteredDays)
                        {
                            case "Saturday" when startTimeParse > new TimeSpan(20, 0, 0): // 20:00 or 8pm
                                return "Personal Care - Night Shift - Saturday";
                                break;

                            case "Sunday" when startTimeParse > new TimeSpan(20, 0, 0): // 20:00 or 8pm
                                return "Personal Care - Night Shift - Sunday";
                                break;

                            case "Saturday":
                                return "Personal Care - Saturday";
                                break;

                            case "Sunday":
                                return "Personal Care - Sunday";
                                break;

                            case var _ when startTimeParse > new TimeSpan(20, 0, 0): // 20:00 or 8pm
                                return "Personal Care - Night Shift - Weekday";
                                break;

                            default:
                                return "Personal Care / Home Care - Weekday";
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Private rate Unassigned");
                        return "";
                    }
                }
                if (funder == "DVA")
                {
                    if (blueCrossService == "Personal Care")
                    {
                        return "DVA - Personal care";
                    }
                    else if (lookoutService.Contains("Sleepover") && lookoutService.Contains("active"))
                    {
                        return "DVA - Overnight personal care active";
                    }
                    else if (lookoutService.Contains("Sleepover") && lookoutService.Contains("inactive"))
                    {
                        return "DVA - Overnight personal care inactive";
                    }
                    else if (lookoutService == "Registered Nurse")
                    {
                        return "DVA - Clinical care";
                    }
                    else
                    {
                        Console.WriteLine("cannot match DVA rate");
                        return "";
                    }
                }

                if (funder == "VHC")
                {
                    if (blueCrossService == "Personal Care")
                    {
                        return "VHC - Personal Care";
                    }

                    else if (blueCrossService == "Home Care")
                    {
                        return "VHC - Domestic Assistance";
                    }
                    else if (blueCrossService == "Respite")
                    {
                        return "VHC - In-Home Respite";
                    }
                    else
                    {
                        Console.WriteLine("cannot match VHC rate");
                        return "";
                    }
                }
                return "";
            }

            static string GetReorderedName(string input)
            {
                string[] nameParts = input.Split(new[] { ", " }, StringSplitOptions.None);
                if (nameParts.Length == 2)
                {
                    string lastName = nameParts[0].Trim();
                    string firstName = nameParts[1].Trim();
                    string fullName = firstName + " " + lastName;
                    return fullName.Replace("  ", " ");
                }
                else
                {
                    // Handle invalid input
                    return "Invalid name format";
                }
            }


            static int GetColumnNumberByHeaderName(ExcelWorksheet worksheet, string headerName)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    if (worksheet.Cells[1, col].Text.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    {
                        return col;
                    }
                }
                // Return a default value or handle the case where the header name is not found
                return -1;
            }

            static string GetFunder(string RosteredFunder)
            {
                string funder;

                if (RosteredFunder.ToLower().Contains("hcp"))
                {
                    funder = "HCP";
                }
                else if (RosteredFunder.ToLower().Contains("chsp"))
                {
                    funder = "CHSP";
                }
                else if (RosteredFunder.ToLower().Contains("community nursing"))
                {
                    funder = "DVA";
                }
                else if (RosteredFunder.ToLower().Contains("private"))
                {
                    funder = "Private";
                }
                else if (RosteredFunder.ToLower().Contains("veteran") || RosteredFunder.ToLower().Contains("vhc"))
                {
                    funder = "VHC";
                }
                else
                {
                    funder = "unknown";
                }

                return funder;
            }



            static DateTime getRuleStartDate(DateTime originalStartDate, int periodWeeks)
            {
                DateTime currentDate = DateTime.Now;
                DateTime nearestStartDate = originalStartDate;

                //Get closest future start date by incrimenting original start date in weeks
                while (nearestStartDate <= currentDate)
                {
                    nearestStartDate = nearestStartDate.AddDays(periodWeeks * 7);
                }

                // Find the nearest prior Monday to the upcoming nearestStartDate
                DayOfWeek dayOfWeek = nearestStartDate.DayOfWeek;
                int daysUntilMonday = (dayOfWeek - DayOfWeek.Monday + 7) % 7;
                DateTime previousMonday = nearestStartDate.AddDays(daysUntilMonday * -1);

                return previousMonday;
            }

        }
    }
}
