using AverageSpreadsExcelReport;
using CT;
using CT.Data;
using CT.Logging;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace speard_report
{
    class CheckSession : ILogSource
    {
        public string LogSourceName
        {
            get
            {
                return "CheckSession";
            }
        }

        protected ITable<AverageSpread> m_AverageTable;
        protected TimeSpan[] m_AsianTradingSession;
        protected TimeSpan[] m_LondonTradingSession;
        protected TimeSpan[] m_USTradingSession;
        protected GetDatabaseTable m_data;
        public CheckSession(GetDatabaseTable data)
        {
            //m_data = data;
            Ini iniReader = Ini.ProgramIniFile;
            m_AverageTable = data.AveraGeTable;
            m_AsianTradingSession = ParseSessionTime(iniReader.ReadSetting("Settings", "AsianTradingSession"));
            m_LondonTradingSession = ParseSessionTime(iniReader.ReadSetting("Settings", "LondonTradingSession"));
            m_USTradingSession = ParseSessionTime(iniReader.ReadSetting("Settings", "USTradingSession"));
        }
        public void FillByseasionTime(out IEnumerable<IGrouping<string, AverageSpread>> group1, out IEnumerable<IGrouping<string, AverageSpread>> group2, out IEnumerable<IGrouping<string, AverageSpread>> group3)
        {
            group1 = AvregateBySession(m_AsianTradingSession[0], m_AsianTradingSession[1]);
            group2 = AvregateBySession(m_LondonTradingSession[0], m_LondonTradingSession[1]);
            group3 = AvregateBySession(m_USTradingSession[0], m_USTradingSession[1]);
        }
        private IEnumerable<IGrouping<string, AverageSpread>> AvregateBySession(TimeSpan start, TimeSpan end)
        {
            try
            {
                //ITable<Symbol> tableSybol = m_data.SymbolTable;

                DateTime startTime = CreateWithTimeSpan(start);
                DateTime endTime = CreateWithTimeSpan(end);
                TimeSpan accuracy = TimeSpan.FromMilliseconds(999);
                if (endTime < startTime)
                {
                    // Ending time is less than starting time, subtract one day from starting time
                    startTime = startTime.Subtract(TimeSpan.FromDays(1));
                }
                this.LogInfo("StartTime: {0}, EndTime: {1}", startTime, endTime);
                TimeSpan duration = endTime - startTime;
                var groupSymbol = m_AverageTable.GetStructs(
                      Search.FieldGreaterOrEqual("TimeStamp", startTime.Add(TimeSpan.FromMinutes(1)))
                      & Search.FieldSmallerOrEqual("TimeStamp", endTime.Add(accuracy))
                      & Search.FieldEquals("Duration", 300)).GroupBy(x => x.Symbol.Replace("/", string.Empty));
                this.LogInfo("<cyan>Number of rows: {0}", groupSymbol.Count());

                return groupSymbol;
            }
            catch (Exception)
            {
                this.LogError("<red> Data not found");
                return null;
            }
        }
        public static DateTime CreateWithTimeSpan(TimeSpan timeSpan)
        {
            DateTime current = new DateTime(2018, 07, 13);
            DateTime datetime = current.Subtract(TimeSpan.FromDays(1)) // Previous day
                .Subtract(current.TimeOfDay) // Set the time
                .Add(timeSpan);
            if (datetime > current)
            {
                datetime = datetime.Subtract(TimeSpan.FromDays(1));
            }
            return datetime;
        }
        public static TimeSpan[] ParseSessionTime(string timeString)
        {
            TimeSpan[] sessionTime = new TimeSpan[2];
            string pattern = @"^\s*(?<StartHour>\d{1,2})\s*:(?<StartMinute>\d{1,2})\s*-" +
                @"\s*(?<EndHour>\d{1,2})\s*:\s*(?<EndMinute>\d{1,2})\s*$";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(timeString);

            if (match.Success)
            {
                sessionTime[0] = TimeSpan.FromHours(int.Parse(match.Groups["StartHour"].Value))
                    .Add(TimeSpan.FromMinutes(int.Parse(match.Groups["StartMinute"].Value)));
                sessionTime[1] = TimeSpan.FromHours(int.Parse(match.Groups["EndHour"].Value))
                    .Add(TimeSpan.FromMinutes(int.Parse(match.Groups["EndMinute"].Value)));
            }
            return sessionTime;
        }
    }
}
