using CT;
using CT.Logging;
using System;
using CT.Data;
using AverageSpreadsExcelReport;
using System.Data.Common;

namespace speard_report
{
    class GetDatabaseTable : ILogSource
    {
        public readonly IDatabase hardDatabase;
        public readonly IStorage hardStorage;
        public ITable<AverageSpread> AveraGeTable { get; private set; }
        public ITable<Symbol> SymbolTable { get; private set; }
        private static Ini iniReader = Ini.ProgramIniFile;
        public string LogSourceName => "GetDatabaseTable";

        public GetDatabaseTable()
        {
            ConnectionString connString = iniReader.ReadSetting("Settings", "DatabaseConnection");
           
             hardStorage = Connector.ConnectStorage(connString, DbConnectionOptions.AllowUnsafeConnections | DbConnectionOptions.VerboseLogging); //?
            if(hardStorage != null)
            {
                try
                {
                    this.LogInfo(("+ Opening..."));
                    hardDatabase = hardStorage.GetDatabase(connString.Location, true);
                }
                catch (ArgumentException)
                {
                    hardDatabase = null;
                }
            }
            if(hardDatabase != null && hardStorage != null)
            {
                Ini iniReader = Ini.ProgramIniFile;
                string tableNameAs = iniReader.ReadString("Settings", "AverageSpreadsTable", "AverageSpreads");
                string tableNameSymbol = iniReader.ReadString("Settings", "SymbolTable", "Symbols");

                this.LogInfo("++ Execution...");
                AveraGeTable = hardDatabase.GetTable<AverageSpread>(TableFlags.AllowCreate, tableNameAs);
                SymbolTable = hardDatabase.GetTable<Symbol>(TableFlags.AllowCreate, tableNameSymbol);
                this.LogInfo("+ Connected...");
            } 
            else
            {
                throw new NullReferenceException("Can not establish a connection to database!");
            }
        }
    }
   
}
