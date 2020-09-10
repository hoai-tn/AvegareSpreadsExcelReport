using AverageSpreadsExcelReport;
using CT;
using CT.Logging;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
namespace speard_report
{
    class Program
    {
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
            GetDatabaseTable getDB = new GetDatabaseTable();
            CheckSession checkSession = new CheckSession(getDB);
            checkSession.FillByseasionTime(
                out IEnumerable<IGrouping<string, AverageSpread>> groupSymbol1,
                out IEnumerable<IGrouping<string, AverageSpread>> groupSymbol2, 
                out IEnumerable<IGrouping<string, AverageSpread>> groupSymbol3);
            IEnumerable<IGrouping<string, AverageSpread>> g = groupSymbol1;

        }
    }
}
