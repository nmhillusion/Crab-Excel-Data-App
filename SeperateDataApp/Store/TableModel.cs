using System.Collections.Generic;

namespace SeperateDataApp.Store
{
    class TableModel
    {
        public string tableName { get; set; }
        private readonly List<List<object>> tableData = new();

        public void SetTableData(List<List<object>> data)
        {
            tableData.Clear();
            tableData.AddRange(data);
        }

        public List<object> GetHeader()
        {
            if (0 < tableData.Count)
            {
                return tableData[0];
            }
            else
            {
                return new List<object>();
            }
        }
    }
}
