using CrabExcelDataApp.Model;
using System.Collections.Generic;

namespace CrabExcelDataApp.Store
{
    class TableStore
    {
        private readonly List<TableModel> data = new();
        private static readonly TableStore instance = new();

        private TableStore() { }

        public static TableStore GetInstance()
        {
            return instance;
        }

        public void SetData(List<TableModel> newTables)
        {
            data.Clear();
            data.AddRange(newTables);
        }

        public TableModel GetSheetAt(int sheetIdx)
        {
            if (0 <= sheetIdx && sheetIdx < data.Count)
            {
                return data[sheetIdx];
            }
            else
            {
                return new TableModel();
            }
        }

        public long Count
        {
            get
            {
                return data.Count;
            }
        }
    }
}
