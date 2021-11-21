using ExcelDataReader;
using SeperateDataApp.Model;
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

        public List<TableModel> ReadData(string excelFilePath)
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

                List<TableModel> totalData = new();

                for (int tableIdx = 0; tableIdx < dataTables.Count; ++tableIdx)
                {
                    DataTable dataTable = dataTables[tableIdx];
                    DataRowCollection dataRowCollection = dataTable.Rows;

                    logHelper.Debug("Table Name: " + dataTable.TableName);


                    List<List<object>> tableData = new();

                    for (int rowIdx = 0; rowIdx < dataRowCollection.Count; ++rowIdx)
                    {
                        DataRow dataRow = dataRowCollection[rowIdx];

                        List<object> rowData = new();
                        rowData.AddRange(dataRow.ItemArray);

                        tableData.Add(rowData);
                    }

                    TableModel tableModel = new();
                    tableModel.tableName = dataTable.TableName;
                    tableModel.SetTableData(tableData);

                    totalData.Add(tableModel);
                }

                return totalData;
            }
            catch (Exception ex)
            {
                logHelper.Error(ex);
                return new List<TableModel>();
            }
        }
    }
}
