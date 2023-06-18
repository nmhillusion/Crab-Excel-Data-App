using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service.Logger;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace CrabExcelDataApp.Service
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
                DataSet dataSet = null;

                using (IExcelDataReader reader = ExcelReaderFactory
                    .CreateReader(
                        File.OpenText(excelFilePath).BaseStream
                    ))
                {

                    logHelper.Info("ResultsCount: " + reader.ResultsCount);
                    logHelper.Info("RowCount: " + reader.RowCount);

                    dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        UseColumnDataType = true
                    });
                }

                DataTableCollection dataTables = dataSet.Tables;

                List<TableModel> totalData = new List<TableModel>();

                for (int tableIdx = 0; tableIdx < dataTables.Count; ++tableIdx)
                {
                    DataTable dataTable = dataTables[tableIdx];
                    DataRowCollection dataRowCollection = dataTable.Rows;

                    logHelper.Debug("Table Name: " + dataTable.TableName);


                    List<List<object>> tableData = new List<List<object>>();

                    for (int rowIdx = 0; rowIdx < dataRowCollection.Count; ++rowIdx)
                    {
                        DataRow dataRow = dataRowCollection[rowIdx];

                        List<object> rowData = new List<object>();
                        rowData.AddRange(dataRow.ItemArray);
                        Debug.WriteLine(this.GetType().Name + " - read data from excel: " + string.Join(", ", rowData));

                        tableData.Add(rowData);
                    }

                    TableModel tableModel = new TableModel();
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
