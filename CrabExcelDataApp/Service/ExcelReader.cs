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
                        if (dataRow.IsNull(0))
                        {
                            continue;
                        }
                        List<object> rowData = new List<object>();
                        rowData.AddRange(dataRow.ItemArray);
                        logHelper.Info(" read data from excel: " + string.Join(", ", rowData));

                        tableData.Add(rowData);
                    }

                    TableModel tableModel = new TableModel
                    {
                        tableName = dataTable.TableName
                    };
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
                logHelper.Info($"sheetCount: {sheetCount}");


                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    int lastRow = 1;
                    int lastColumn = 1;

                    logHelper.Info($"lastRow: {lastRow}");
                    logHelper.Info($"lastColumn: {lastColumn}");

                    List<List<object>> sheetData = new List<List<object>>();

                    for (int rowNum = 1; rowNum <= lastRow; ++rowNum)
                    {
                        if (null == worksheet.Cells[rowNum, 1].Value)
                        {
                            logHelper.Warn($"Ignore row #{rowNum}");
                            continue;
                        }

                        Range row_ = worksheet.Rows[rowNum];

                        if (!row_.Hidden)
                        {
                            List<object> rowData = new List<object>();
                            for (int colNum = 1; colNum <= lastColumn; ++colNum)
                            {
                                Range cell_ = worksheet.Cells[rowNum, colNum];
                                rowData.Add(cell_.Value);
                            }
                            logHelper.Info("read data from excel: " + string.Join(", ", rowData));
                            sheetData.Add(rowData);
                        }
                    }

                    TableModel model = new TableModel()
                    {
                        tableName = worksheet.Name
                    };
                    model.SetTableData(sheetData);
                    totalData.Add(model);

                    Marshal.ReleaseComObject(worksheet);
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

        private Range getFirstDataCell(Worksheet worksheet)
        {
            Range firstCell = worksheet.Cells[1, 1];

            int lastRow = firstCell.End[XlDirection.xlDown].Row;
            int lastColumn = firstCell.End[XlDirection.xlToRight].Column;
            // TODO: impl
            return null;
        }
    }
}
