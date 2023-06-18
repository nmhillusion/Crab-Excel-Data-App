using CrabExcelDataApp.Util;
using System.Collections.Generic;
using System.Linq;

namespace CrabExcelDataApp.Model
{
    class TableModel
    {
        public string tableName;
        public int SIZE_OF_HEADER_TO_GET = 1;
        private readonly List<List<object>> tableData = new List<List<object>>();

        public void SetTableData(List<List<object>> data)
        {
            tableData.Clear();
            tableData.AddRange(data);
        }

        public List<object> GetHeader()
        {
            if (0 < tableData.Count)
            {
                List<object> headers = new List<object>();
                var firstRowRange = tableData.GetRange(0, SIZE_OF_HEADER_TO_GET);

                if (CollectionUtil.IsNullOrEmpty(firstRowRange))
                {
                    return new List<object>();
                }

                headers.AddRange(firstRowRange.ElementAt(0));

                return headers;
            }
            else
            {
                return new List<object>();
            }
        }

        public List<List<object>> GetBody()
        {
            return 1 < tableData.Count ? tableData.GetRange(1, tableData.Count - 1) : new List<List<object>>();
        }

        public List<object> GetDataAtColumnIdx(int columnIdxToGet)
        {
            List<object> columnData = new List<object>();

            List<List<object>> bodyData = GetBody();
            if (0 <= columnIdxToGet && columnIdxToGet < GetHeader().Count)
            {
                foreach (List<object> row in bodyData)
                {
                    columnData.Add(
                        row[columnIdxToGet]
                    );
                }
            }

            return columnData;
        }
    }
}
