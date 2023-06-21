using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service.Logger;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

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
            return ReadData(excelFilePath, true);
        }

        public List<TableModel> ReadData(string excelFilePath, bool isIgnoreHiddenRows)
        {
            if (!isIgnoreHiddenRows)
            {
                return ReadDataAllRows(excelFilePath);
            }
            else
            {
                return ReadDataAndIgnoreHiddenRows(excelFilePath);
            }
        }

        public List<TableModel> ReadDataAllRows(string excelFilePath)
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
                    System.Data.DataTable dataTable = dataTables[tableIdx];
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

        public List<TableModel> ReadDataAndIgnoreHiddenRows(string excelFilePath)
        {
            logHelper.Info("Read Excel at " + excelFilePath);

            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                excelApp = new Application();
                workbook = excelApp.Workbooks.Open(excelFilePath);

                List<TableModel> totalData = new List<TableModel>();

                int sheetCount = workbook.Worksheets.Count;
                logHelper.Info($" sheetCount: {sheetCount}");


                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    Range firstCell = worksheet.Cells[1, 1];
                    Range lastCellInColumn = firstCell.End[XlDirection.xlDown];
                    Range lastCellInRow = firstCell.End[XlDirection.xlToRight];

                    int lastRow = lastCellInColumn.Row;
                    int lastColumn = lastCellInRow.Column;
                    logHelper.Info($" lastRow: {lastRow}");
                    logHelper.Info($" lastColumn: {lastColumn}");

                    List<List<object>> sheetData = new List<List<object>>();

                    for (int rowNum = 1; rowNum <= lastRow; ++rowNum)
                    {
                        Range row_ = worksheet.Rows[rowNum];

                        if (!row_.Hidden)
                        {
                            List<object> rowData = new List<object>();
                            for (int colNum = 1; colNum <= lastColumn; ++colNum)
                            {
                                Range cell_ = worksheet.Cells[rowNum, colNum];
                                rowData.Add(cell_.Value);
                            }
                            sheetData.Add(rowData);
                        }
                    }

                    TableModel model = new TableModel()
                    {
                        tableName = worksheet.Name
                    };
                    model.SetTableData(sheetData);
                    totalData.Add(model);
                }

                return totalData;
            }
            catch (Exception ex)
            {
                logHelper.Error(ex);
                return new List<TableModel>();
            }
            finally
            {
                if (null != workbook)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }

                if (null != excelApp)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }
}
