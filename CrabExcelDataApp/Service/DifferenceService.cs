using CrabExcelDataApp.Model;
using System.Collections.Generic;

namespace CrabExcelDataApp.Service
{
    class DifferenceService
    {
        public ISet<string> DistinctListObject(IList<object> allValues)
        {
            SortedSet<string> sortedSet = new();

            foreach (object item in allValues)
            {
                sortedSet.Add(item.ToString());
            }

            return sortedSet;
        }

        public List<List<object>> FilterData(TableModel tableModel, int columnIdxToCompare, string targetValueToCompare)
        {
            List<List<object>> filteredData = new();

            List<List<object>> bodyTable = tableModel.GetBody();

            foreach (List<object> rowData in bodyTable)
            {
                object currentCellToCompare = rowData[columnIdxToCompare];
                if (0 ==
                    string.Compare(
                        targetValueToCompare.Trim(),
                        currentCellToCompare.ToString().Trim(),
                        System.StringComparison.Ordinal)
                )
                {
                    filteredData.Add(rowData);
                }
            }

            return filteredData;
        }
    }
}
