using AverageSpreadsExcelReport;
using CT;
using CT.Data;
using CT.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace speard_report
{
    class BestAverage
    {
        private List<ITable<LiveQuote>> m_ListTables;
        public BestAverage(GetDatabaseTable data, string[] outputTables)
        { 
            m_ListTables = new List<ITable<LiveQuote>>(); 
            foreach(string outputTable in outputTables)
            {
                try
                {
                    m_ListTables.Add(data.hardDatabase.GetTable<LiveQuote>(TableFlags.AllowCreate, outputTable)); // create table ini in program
                } catch(Exception ex)
                {
                    this.LogError("Cannot get {0}. Error details: {1}", outputTable, ex.Message);
                }
                this.LogInfo("Output tables: <cyan>{0}", string.Join(", ", outputTables));
            }
        }
        public void Run(List<Average> listAvg)
        {
            var dicAvgGroup = CreateGroupBestAvg(listAvg);// format list avg type dictionary
            FilterBrokerName(dicAvgGroup);
        }
        private Dictionary<SymbolBrokerName, double> CreateGroupBestAvg(List<Average> listAvg) // 
        {
            Dictionary<SymbolBrokerName, double> groupBestAvg = new Dictionary<SymbolBrokerName, double>();
            foreach(var itemAvg in listAvg) // loop total list avg from class filterDataTable
            {
                SymbolBrokerName key = new SymbolBrokerName //create key for dictionary
                {
                    Symbol = itemAvg.Symbol,
                    BrokerName = itemAvg.Broker,
                };
                if (!groupBestAvg.ContainsKey(key))
                    groupBestAvg.Add(key, (double)itemAvg.Avg); // add value in dic
                else
                    groupBestAvg[key] = (double)itemAvg.Avg; // update 
            }
            return groupBestAvg;
        }
        private void FilterBrokerName(Dictionary<SymbolBrokerName, double> listBestAvg)
        {
            Ini iniReader = Ini.ProgramIniFile;
            string[] GBEBrokers = iniReader.ReadSection("GBEBrokers");
            string brokerName = null;
            for(int i = 0; i < m_ListTables.Count(); i++) // loop data from the gbe table 
            {   
                try
                {
                    brokerName = GBEBrokers[i];
                }
                catch (IndexOutOfRangeException ex)
                {
                    this.LogError(ex, "Not found table to update!");
                    continue;
                }
                LiveQuote[] rows = m_ListTables[i].GetStructs().ToArray();// get three table 
                if(rows.Length > 0)
                {
                    this.LogInfo("Find the best average spread for broker {0}", brokerName);
                    var listBestAvgByBroker = listBestAvg.Where(x => x.Key.BrokerName == brokerName ); // filter by broker name same brokername of the ini file
                    UpdateBestAvgTable(rows, listBestAvgByBroker, m_ListTables[i]);//update data
                }
            }
        }
        public void UpdateBestAvgTable(LiveQuote[] rows, IEnumerable<KeyValuePair<SymbolBrokerName, double>> listBestAvgByBroker, ITable<LiveQuote> liveQuoteTable)
        {
            foreach(var row in rows)// loop row in table 
            {
                var rowUpdate = row;
                foreach (var itemBestAvg in listBestAvgByBroker.Where(x => x.Key.Symbol == row.Symbol)) // loop item avg with conditions same a symbol
                {
                    rowUpdate.BrokerName = itemBestAvg.Key.BrokerName; // update column broker name
                    rowUpdate.SpreadAvg = Math.Round(itemBestAvg.Value, 2, MidpointRounding.AwayFromZero); 
                    liveQuoteTable.Update(rowUpdate);
                    this.LogInfo($"Best average spread for broker {rowUpdate.BrokerName} and symbol {rowUpdate.Symbol} is {itemBestAvg.Value:N}");
                }
            }
        }

    }
}
