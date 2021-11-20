using System.Collections.Generic;

namespace SeperateDataApp.Store
{
    class TableStore
    {
        private readonly List<List<List<string>>> data = new();
        private static readonly TableStore instance = new();

        private TableStore() { }

        public static TableStore GetInstance()
        {
            return instance;
        }

        public void SetData(List<List<List<string>>> newData)
        {
            data.Clear();
            data.AddRange(newData);
        }

        public List<List<string>> GetSheetAt(int sheetIdx)
        {
            if (0 <= sheetIdx && sheetIdx < data.Count)
            {
                return data[sheetIdx];
            }
            else
            {
                return new List<List<string>>();
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
