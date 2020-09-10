using AverageSpreadsExcelReport;
using CT;
using CT.Data;
using CT.Logging;
using System.Collections.Generic;
using System.Linq;

namespace speard_report
{
    class FilterDataTable
    {

        private List<Average> AverageGroup(List<IGrouping<string, AverageSpread>> GroupSymbol)
        {

            List<Average> listAvg = new List<Average>();
            foreach (var itemSymbol in GroupSymbol)
            {
                var groupBroke = itemSymbol.Select(x => x).GroupBy(x => x.BrokerName).ToList();
                foreach (var itemBroke in groupBroke)
                {

                }
            }
            return listAvg;
        }

    }
}
