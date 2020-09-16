using AverageSpreadsExcelReport;
using CT;
using CT.Data;
using System.Collections.Generic;
using System.Linq;

namespace speard_report
{
    class FilterDataTable
    {
        GetDatabaseTable m_data;
        public string[] BrokerNames;
        public string[] OtherNames;
        public IEnumerable<IGrouping<string, AverageSpread>> GroupSymbol;
        public List<Average> ListAverages;
        public Dictionary<string, double?[]> m_ListBroker;
        public Dictionary<string, double?[]> m_ListOtherBroker;
        public FilterDataTable(GetDatabaseTable data, List<IGrouping<string, AverageSpread>> symbols)
        {
            m_data = data;
            Ini iniReader = Ini.ProgramIniFile;
            GroupSymbol = symbols;
            ListAverages = AverageGroup(symbols.ToList());
            BrokerNames = CheckBrokerName(iniReader.ReadSection("GBEBrokers"));

            if (iniReader.ReadSection("OtherBrokers").Length <= 0)
            {
                OtherNames = ListAverages.Select(x => x.Broker).Distinct().OrderBy(x => x).Except(BrokerNames).Concat(BrokerNames).ToArray();
            }
            else
            {
                OtherNames = iniReader.ReadSection("OtherBrokers");
            }
            m_ListBroker = GetDicionary(ListAverages, BrokerNames);
            m_ListOtherBroker = GetDicionary(ListAverages, OtherNames);
        }
        public List<Average> AverageGroup(List<IGrouping<string, AverageSpread>> GroupSymbol)
        {
            List<Average> listAvg = new List<Average>();
            foreach (var itemSymbol in GroupSymbol)
            {
                var brokers = itemSymbol.GroupBy(x => x.BrokerName).ToList();//group broker
                foreach (var broker in brokers)
                {
                    Average average = new Average();
                    var valueAvg = broker.Average(x => x.AvgSpread);// averaged a broker
                    average.Avg = valueAvg;
                    average.Broker = broker.Key;
                    average.Symbol = itemSymbol.Key;
                    listAvg.Add(average);
                }
            }
            return SortSymbol(listAvg);
        }
        private List<Average> SortSymbol(List<Average> lst)
        {
            List<Average> lstSort = new List<Average>();
            ITable<Symbol> table = m_data.SymbolTable;
            string[] sybolName = table.GetStructs(Search.None, ResultOption.SortAscending(nameof(Symbol.ID))).Select(x => x.currencypairname.Replace("/", string.Empty)).ToArray();
            foreach (var item in sybolName)
            {
                var sortSymbol = lst.Where(x => x.Symbol == item); 
                foreach (var y in sortSymbol)
                {
                    lstSort.Add(new Average 
                    {
                        Symbol = y.Symbol,
                        Avg = y.Avg,
                        Broker = y.Broker,
                    });
                }
            }
            return lstSort;
        }
        private string[] CheckBrokerName(string[] listBroker)
        {
            if (listBroker.Length <= 0)
                return ListAverages.Select(x => x.Broker).Distinct().OrderBy(x => x).ToArray();
            return listBroker;

        }
        public Dictionary<string, double?[]> GetDicionary(List<Average> averages, string[] arrayBorker)// create [, ] inside 
        {
            Dictionary<string, double?[]> rows = new Dictionary<string, double?[]>();
            string[] grpNames = arrayBorker.Distinct().ToArray();
            var filter = averages.Where(x => grpNames.Contains(x.Broker)).ToList();

            int count = grpNames.Length;
            foreach (var a in filter)
            {  //neu chua co key trong rows
                if (!rows.ContainsKey(a.Symbol))
                {
                    rows.Add(a.Symbol, new double?[count]);
                }
                // them value theo tri tu  0 -1
                for (int i = 0; i < count; i++)
                {
                    if (a.Broker == grpNames[i])// if borker just create === item 
                    {
                        rows[a.Symbol][i] = a.Avg; //row'll insert value avg
                        break;
                    }
                }
            }
            return rows;
        }
    }
}
