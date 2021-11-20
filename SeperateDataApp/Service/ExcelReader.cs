using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SeperateDataApp.Service
{
    class ExcelReader
    {
        private readonly LogHelper logHelper;

        public ExcelReader()
        {
            logHelper = new LogHelper(this);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public List<List<List<string>>> ReadData(string excelFilePath)
        {
            logHelper.Info("Read Excel at " + excelFilePath);
            try
            {
                using IExcelDataReader reader = ExcelReaderFactory
                    .CreateReader(
                        File.OpenText(excelFilePath).BaseStream
                    );

                logHelper.Info("ResultsCount: " + reader.ResultsCount);
                logHelper.Info("RowCount: " + reader.RowCount);

                DataSet dataSet = reader.AsDataSet();
                DataTableCollection dataTables = dataSet.Tables;

                List<List<List<string>>> totalData = new List<List<List<string>>>();

                for (int tableIdx = 0; tableIdx < dataTables.Count; ++tableIdx)
                {
                    DataTable dataTable = dataTables[tableIdx];
                    DataRowCollection dataRowCollection = dataTable.Rows;

                    List<List<string>> tableData = new List<List<string>>();

                    for (int rowIdx = 0; rowIdx < dataRowCollection.Count; ++rowIdx)
                    {
                        DataRow dataRow = dataRowCollection[rowIdx];
                        object[] cells = dataRow.ItemArray;

                        List<string> rowData = new List<string>();

                        for (int colIdx = 0; colIdx < cells.Length; ++colIdx)
                        {
                            rowData.Add(cells[colIdx].ToString());
                        }

                        tableData.Add(rowData);
                    }

                    totalData.Add(tableData);
                }

                return totalData;
            }
            catch (Exception ex)
            {
                logHelper.Error(ex);
                return new List<List<List<string>>>();
            }
        }
    }
}
