using AverageSpreadsExcelReport;
using CT;
using CT.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
namespace speard_report
{
    class Program
    {
        private static TimeSpan timesPeriod = TimeSpan.FromDays(1);
        static void Main(string[] args)
        {
            new MySql.Data.MySqlClient.MySqlConnection().Dispose();
            var logFile = LogFile.CreateProgramLogFile(LogFileFlags.UseDateTimeFileName);
            logFile.Level = LogLevel.Debug;
            LogConsole logConsole = LogConsole.Create(LogConsoleFlags.DefaultShort, LogLevel.Debug);
            foreach (var arg in args)
            {
                CT.Console.SystemConsole.Title += arg + " ";
            }
            try
            {
                var task = Task.Factory.StartNew(CreateReport);

                DateTime lastExitPressed = DateTime.MinValue;
                while (true)
                {
                    if (Console.KeyAvailable)
                    {
                        var cki = Console.ReadKey(true);
                        var key = cki.Key;
                        var modifier = cki.Modifiers;
                        if (key == ConsoleKey.Q && modifier == ConsoleModifiers.Control)
                        {
                            if (lastExitPressed.AddSeconds(1) >= DateTime.Now)
                            {
                                Logger.LogInfo("Program", "User pressed CTRL-Q-Q, program exit.");
                                break;
                            }
                            lastExitPressed = DateTime.Now;
                            Logger.LogInfo("Program", "Press CRT-Q again to exit.");
                        }
                    }
                    Thread.Sleep(200);
                }

                if (task.IsFaulted)
                {
                    Logger.LogError(task.GetType().ToString(), task.Exception, task.Exception.ToText());
                }
            }
            finally
            {
                logFile.Close();
                logConsole.Close();
            }
        }
        private static void CreateReport()
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            GetDatabaseTable getDB = new GetDatabaseTable();
            CheckSession checkSession = new CheckSession(getDB);
            checkSession.FillByseasionTime(
                out IEnumerable<IGrouping<string, AverageSpread>> groupSymbolAsia,
                out IEnumerable<IGrouping<string, AverageSpread>> groupSymbolLondon, 
                out IEnumerable<IGrouping<string, AverageSpread>> groupSymbolUS);
            if (groupSymbolAsia.Count() != 0 || groupSymbolLondon.Count() != 0 || groupSymbolUS.Count() != 0)
            {
                FilterDataTable filterDataAsia = new FilterDataTable(getDB, groupSymbolAsia.ToList());
                FilterDataTable filterDataLondon = new FilterDataTable(getDB, groupSymbolLondon.ToList());
                FilterDataTable filterDataUS = new FilterDataTable(getDB, groupSymbolUS.ToList());
                using (ExcelPackage excel = new ExcelPackage())
                { // tao sheet
                    var LondonTradingSession = excel.Workbook.Worksheets.Add("London Session");
                    var USTradingSession = excel.Workbook.Worksheets.Add("New York Session");
                    var AsianTradingSession = excel.Workbook.Worksheets.Add("Asia Session");
                    // create table sheet Asia
                    new WriteExcel(
                        AsianTradingSession, 
                        filterDataAsia.BrokerNames, 
                        filterDataAsia.OtherNames, 
                        filterDataAsia.m_ListBroker, 
                        filterDataAsia.m_ListOtherBroker);
                    Logger.LogInfo("Main", "Write complete session Asia.");
                    // create table sheet London
                    new WriteExcel(
                        LondonTradingSession,
                        filterDataLondon.BrokerNames,
                        filterDataLondon.OtherNames,
                        filterDataLondon.m_ListBroker,
                        filterDataLondon.m_ListOtherBroker);
                    Logger.LogInfo("Main", "Write complete session London.");
                    // create sheet table US
                    new WriteExcel(
                        USTradingSession,
                        filterDataUS.BrokerNames,
                        filterDataUS.OtherNames,
                        filterDataUS.m_ListBroker,
                        filterDataUS.m_ListOtherBroker);
                    Logger.LogInfo("Main", "Write complete session USA.");
                    //luu duong dan
                    string tPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string path = Directory.GetCurrentDirectory() + "\\reports";
                    Directory.CreateDirectory(path);
                    string fileName = path + string.Format("\\AverageSpreadsReport_{0}.xlsx", DateTime.UtcNow.ToString("yyyy.MM.dd"));
                    FileInfo excelFile = new FileInfo(fileName);
                    excel.SaveAs(excelFile);

                    Logger.LogInfo("Main", $"Save file as {fileName}.");
                    // gui mail
                    Ini iniReader = Ini.ProgramIniFile;
                    Mailer mailer = new Mailer(iniReader);
                    Logger.LogInfo("Main", "<cyan>Sending an email with excel report...");
                    mailer.SendMail("Average Spreads Report", "<h3>Average Spreads Report</h3>", "Average Spreads Report", new List<string> { fileName });
                    Logger.LogInfo("Main", "Sending email is successfully.");

                    // statistical best average
                    string[] outTable = iniReader.ReadSection("OutputTables");
                    var totalListAvg = filterDataAsia.ListAverages.Concat(filterDataLondon.ListAverages).Concat(filterDataUS.ListAverages).ToList();
                    if(outTable.Count() > 0)
                    {
                        BestAverage b = new BestAverage(getDB, outTable);
                        b.Run(totalListAvg);
                    }
                    else
                    {
                        Logger.LogError("Main", "Not found table live quote to compare!");
                    }

                }
            } else
            {
                Logger.LogWarning("Main", "Not found data Today!");
            }
            Logger.LogInfo("Main", $"Program will start after {timesPeriod.FormatTime()}.");

            Logger.LogInfo("Main", watch.Elapsed.ToString());
            watch.Stop();
        }
    }
}