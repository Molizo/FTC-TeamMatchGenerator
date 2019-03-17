using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace FTC_TeamMatchGenerator
{
    public class Match
    {
        public string MatchID {get;set;}
        public string RedTeam1 {get;set;}
        public string RedTeam2 {get;set;}
        public string BlueTeam1 {get;set;}
        public string BlueTeam2 {get;set;}
    }

    public class Team
    {
        public string TeamID {get;set;}
        public string TeamName {get;set;}
    }

    internal class Program
    {
        private static List<Match> matches = new List<Match>();
        private static List<Team> teams = new List<Team>();
        private static string providedBy;

        static void Main(string[] args)
        {
            Console.Title = "FTC Individual Team Match List generator";
            Console.WriteLine("+------------------------------------------+");
            Console.WriteLine("| FTC Individual team match list generator |");
            Console.WriteLine("|                                          |");
            Console.WriteLine("|   Copyright Mihnea Theodor Visoiu 2019   |");
            Console.WriteLine("+------------------------------------------+");
            Console.WriteLine();
            Console.WriteLine();
            Console.Write("Please enter text for the \"Provided by\" field: ");
            providedBy = Console.ReadLine();
            Console.Write("Is the file ready for processing? (YES/NO)    ");
            string input = Console.ReadLine();
            Console.WriteLine();
            Console.WriteLine();
            if (input.ToLower().Contains('n'))
            {
                Console.WriteLine("The program has terminated...");
                Console.WriteLine("Press any key to continue");
                Console.ReadKey();
                return;
            }
            
            Console.WriteLine("Reading data");
            ReadMatches();
            ReadTeams();
            
            GenerateMatchSheets();

            Console.WriteLine("The program has terminated...");
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
            return;
        }

        static void GenerateMatchSheets()
        {
            Console.WriteLine("Generating match sheets for each team");
            try
            {
                File.Delete(@".\output.xlsx");
                FileInfo fi = new FileInfo(@".\output.xlsx");
                using (ExcelPackage workbook = new ExcelPackage(fi))
                {
                    foreach(var team in teams)
                    {
                        Console.WriteLine("Processing team # " + team.TeamID);
                        workbook.Workbook.Worksheets.Add(team.TeamID);
                        ExcelWorksheet teamWorksheet = workbook.Workbook.Worksheets[team.TeamID];

                        teamWorksheet.Name = team.TeamID;
                        teamWorksheet.Cells["A1:D1"].Merge = true;
                        teamWorksheet.Cells["A2:D2"].Merge = true;

                        teamWorksheet.Cells["A2:D2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        teamWorksheet.Cells["A1:D1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        teamWorksheet.Cells["A1:D2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        teamWorksheet.Cells["A1:D2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;

                        teamWorksheet.PrinterSettings.FitToWidth = 0;
                        teamWorksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                        teamWorksheet.PrinterSettings.PaperSize = ePaperSize.A5;
                        teamWorksheet.PrinterSettings.BottomMargin = 0m;
                        teamWorksheet.PrinterSettings.TopMargin = 0m;
                        teamWorksheet.PrinterSettings.LeftMargin = 0m;
                        teamWorksheet.PrinterSettings.RightMargin = 0m;

                        var cell = teamWorksheet.Cells[1,1];
                        cell.Value = providedBy;
                        cell.Style.Font.Color.SetColor(255,117,113,113);
                        cell.Style.Font.Size = 14;

                        cell = teamWorksheet.Cells[2,1];
                        cell.Value = "Team #" + team.TeamID + " | " + team.TeamName;
                        cell.Style.Font.Size = 17;

                        cell = teamWorksheet.Cells[3,1];
                        cell.Value = "Match";
                        cell.Style.Font.Size = 14;

                        cell = teamWorksheet.Cells[3,2];
                        cell.Value = "Partner";
                        cell.Style.Font.Size = 14;

                        cell = teamWorksheet.Cells[3,3];
                        cell.Value = "Opponent 1";
                        cell.Style.Font.Size = 14;
                        
                        cell=teamWorksheet.Cells[3,4];
                        cell.Value = "Opponent 2";
                        cell.Style.Font.Size = 14;

                        teamWorksheet.Cells["A3:D3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        teamWorksheet.Cells["A3:D3"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        teamWorksheet.Cells["A3:D3"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        teamWorksheet.Cells["A3:D3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        int currentRow = 4;
                        foreach(var match in matches)
                        {
                            if(match.RedTeam1.TrimEnd('*') == team.TeamID || match.RedTeam2.TrimEnd('*') == team.TeamID || match.BlueTeam1.TrimEnd('*') == team.TeamID || match.BlueTeam2.TrimEnd('*') == team.TeamID)
                            {
                                Console.WriteLine("Found match #"+match.MatchID);

                                cell = teamWorksheet.Cells[currentRow,1];
                                cell.Value = match.MatchID;
                                if((match.RedTeam1.Contains('*') && match.RedTeam1.TrimEnd('*') == team.TeamID) ||
                                   (match.RedTeam2.Contains('*') && match.RedTeam2.TrimEnd('*') == team.TeamID) ||
                                   (match.BlueTeam1.Contains('*') && match.BlueTeam1.TrimEnd('*') == team.TeamID) ||
                                   (match.BlueTeam2.Contains('*') && match.BlueTeam2.TrimEnd('*') == team.TeamID))
                                    cell.Value = match.MatchID + "*";
                                cell.Style.Font.Size = 14;

                                if(match.RedTeam1.TrimEnd('*') == team.TeamID)
                                {
                                    cell.Style.Font.Color.SetColor(Color.Red);
                                    cell = teamWorksheet.Cells[currentRow+1,1];
                                    cell.Value="Red";
                                    cell.Style.Font.Color.SetColor(Color.Red);
                                    cell.Style.Font.Size = 14;
                                    foreach(var team1 in teams)
                                    {
                                        if(team1.TeamID == match.RedTeam2.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,2];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,2];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.BlueTeam1.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,3];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,3];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.BlueTeam2.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,4];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,4];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                    }
                                }

                                else if(match.RedTeam2.TrimEnd('*') == team.TeamID)
                                {
                                    cell.Style.Font.Color.SetColor(Color.Red);
                                    cell = teamWorksheet.Cells[currentRow+1,1];
                                    cell.Value="Red";
                                    cell.Style.Font.Color.SetColor(Color.Red);
                                    cell.Style.Font.Size = 14;
                                    foreach(var team1 in teams)
                                    {
                                        if(team1.TeamID == match.RedTeam1.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,2];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,2];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.BlueTeam1.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,3];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,3];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.BlueTeam2.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,4];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,4];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                    }
                                }
                                else if(match.BlueTeam1.TrimEnd('*') == team.TeamID)
                                {
                                    cell.Style.Font.Color.SetColor(Color.Blue);
                                    cell = teamWorksheet.Cells[currentRow+1,1];
                                    cell.Value="Blue";
                                    cell.Style.Font.Color.SetColor(Color.Blue);
                                    cell.Style.Font.Size = 14;
                                    foreach(var team1 in teams)
                                    {
                                        if(team1.TeamID == match.BlueTeam2.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,2];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,2];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.RedTeam1.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,3];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,3];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.RedTeam2.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,4];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,4];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                    }
                                }

                                else
                                {
                                    cell.Style.Font.Color.SetColor(Color.Blue);
                                    cell = teamWorksheet.Cells[currentRow+1,1];
                                    cell.Value="Blue";
                                    cell.Style.Font.Color.SetColor(Color.Blue);
                                    cell.Style.Font.Size = 14;
                                    foreach(var team1 in teams)
                                    {
                                        if(team1.TeamID == match.BlueTeam1.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,2];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,2];
                                            cell.Style.Font.Color.SetColor(Color.Blue);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.RedTeam1.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,3];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,3];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                        else if(team1.TeamID == match.RedTeam2.TrimEnd('*'))
                                        {
                                            cell = teamWorksheet.Cells[currentRow,4];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamID;
                                            cell.Style.Font.Size = 17;

                                            cell = teamWorksheet.Cells[currentRow+1,4];
                                            cell.Style.Font.Color.SetColor(Color.Red);
                                            cell.Value = team1.TeamName;
                                            cell.Style.Font.Size = 14;
                                        }
                                    }
                                }

                                teamWorksheet.Cells[currentRow+1,1,currentRow+1,4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                teamWorksheet.Cells[currentRow,1,currentRow,4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                teamWorksheet.Cells[currentRow,1,currentRow+1,4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                teamWorksheet.Cells[currentRow,1,currentRow+1,4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                currentRow = currentRow+2;
                            }
                        }
                        teamWorksheet.Column(1).Width = 8;
                        teamWorksheet.Column(2).Width = 31;
                        teamWorksheet.Column(3).Width = 31;
                        teamWorksheet.Column(4).Width = 31;
                        teamWorksheet.Cells[teamWorksheet.Dimension.Address].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                    workbook.Save();
                }
            }
            catch
            {
                Console.WriteLine("Error writing to file!");
            }
        }

        static void ReadMatches()
        {
            Console.WriteLine("Reading matches");
            try
            {
                FileInfo fi = new FileInfo(@".\input.xlsx");
                using (ExcelPackage workbook = new ExcelPackage(fi))
                {
                    ExcelWorksheet worksheet = workbook.Workbook.Worksheets["Matches"];
                    int currentRow = 2;
                    while (true)
                    {
                        Match inputMatch = new Match();
                        if (worksheet.Cells[currentRow, 1].Text == String.Empty)
                            break;
                        inputMatch.MatchID = worksheet.Cells[currentRow, 1].Text;
                        inputMatch.RedTeam1 = worksheet.Cells[currentRow, 2].Text;
                        inputMatch.RedTeam2 = worksheet.Cells[currentRow, 3].Text;
                        inputMatch.BlueTeam1 = worksheet.Cells[currentRow, 4].Text;
                        inputMatch.BlueTeam2 = worksheet.Cells[currentRow, 5].Text;
                        matches.Add(inputMatch);
                        Console.WriteLine("Added match: " + 
                            inputMatch.MatchID + " " +
                            inputMatch.RedTeam1 + " " +
                            inputMatch.RedTeam2 + " " +
                            inputMatch.BlueTeam1 + " " + 
                            inputMatch.BlueTeam2);
                        currentRow++;
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error reading file!");
            }
        }
        static void ReadTeams()
        {
            Console.WriteLine("Reading teams");
            try
            {
                FileInfo fi = new FileInfo(@".\input.xlsx");
                using (ExcelPackage workbook = new ExcelPackage(fi))
                {
                    ExcelWorksheet worksheet = workbook.Workbook.Worksheets["Teams"];
                    int currentRow = 2;
                    while (true)
                    {
                        Team inputTeam = new Team();
                        if (worksheet.Cells[currentRow, 1].Text == String.Empty)
                            break;
                        inputTeam.TeamID = worksheet.Cells[currentRow, 1].Text;
                        inputTeam.TeamName = worksheet.Cells[currentRow, 2].Text;
                        if(inputTeam.TeamName.Length > 18)
                            inputTeam.TeamName = inputTeam.TeamName.Substring(0,18);
                        teams.Add(inputTeam);
                        Console.WriteLine("Added team: " + 
                            inputTeam.TeamID + " " +
                            inputTeam.TeamName);
                        currentRow++;
                    }
                }
            }
            catch
            {
                Console.WriteLine("Error reading file!");
            }
        }
    }
}
