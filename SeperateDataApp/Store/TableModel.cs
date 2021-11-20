﻿using System.Collections.Generic;

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

        public List<List<object>> GetBody()
        {
            return 1 < tableData.Count ? tableData.GetRange(1, tableData.Count - 1) : new List<List<object>>();
        }

        public List<object> GetDataAtColumnIdx(int columnIdxToGet)
        {
            List<object> columnData = new();

            List<List<object>> bodyData = GetBody();
            if (0 <= columnIdxToGet && columnIdxToGet < bodyData.Count)
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
