using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service.Logger;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
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
            return ReadData(excelFilePath, new ExcelFilterModel()
            {
                isStandardTemplate = true,
                isFilterIgnoreHiddenRows = false,
                startRowNum = 1,
            });
        }

        public List<TableModel> ReadData(string excelFilePath, ExcelFilterModel excelFilterModel)
        {
            if (excelFilterModel.isStandardTemplate && !excelFilterModel.isFilterIgnoreHiddenRows)
            {
                return ReadDataAllRows(excelFilePath, excelFilterModel);
            }
            else
            {
                return ReadDataWithAdvantageFilters(excelFilePath, excelFilterModel);
            }
        }

        public List<TableModel> ReadDataAllRows(string excelFilePath, ExcelFilterModel excelFilterModel)
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
                    int startRowIndex = Math.Max(0, excelFilterModel.startRowNum - 1);

                    logHelper.Debug("Table Name: " + dataTable.TableName);


                    List<List<object>> tableData = new List<List<object>>();

                    for (int rowIdx = startRowIndex; rowIdx < dataRowCollection.Count; ++rowIdx)
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

        public List<TableModel> ReadDataWithAdvantageFilters(string excelFilePath, ExcelFilterModel excelFilterModel)
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
                    Range firstCell_ = GetFirstDataCell(worksheet);
                    Range lastCell_ = GetLastDataCell(worksheet, firstCell_);
                    logHelper.Info($"first cell: [{firstCell_.Row}, {firstCell_.Column}]");
                    logHelper.Info($"last cell: [{lastCell_.Row}, {lastCell_.Column}]");

                    logHelper.Info($"lastRow: {lastCell_.Row}");
                    logHelper.Info($"lastColumn: {lastCell_.Column}");

                    List<List<object>> sheetData = new List<List<object>>();
                    int startRowNum = Math.Max(firstCell_.Row, excelFilterModel.startRowNum);

                    for (int rowNum = startRowNum; rowNum <= lastCell_.Row; ++rowNum)
                    {
                        Range row_ = worksheet.Rows[rowNum];

                        if (IsValidHiddenCondition(row_, excelFilterModel))
                        {
                            List<object> rowData = new List<object>();
                            for (int colNum = firstCell_.Column; colNum <= lastCell_.Column; ++colNum)
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
                workbook?.Close();
                excelApp?.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool IsValidHiddenCondition(Range range_, ExcelFilterModel excelFilterModel)
        {
            return !excelFilterModel.isFilterIgnoreHiddenRows || !range_.Hidden;
        }

        private Range GetFirstDataCell(Worksheet worksheet)
        {
            Range firstCell = worksheet.Cells[1, 1];

            if (0 < worksheet.Application.WorksheetFunction.CountA(firstCell))
            {
                return firstCell;
            }

            int lastRow = firstCell.End[XlDirection.xlDown].Row;
            int lastColumn = firstCell.End[XlDirection.xlToRight].Column;

            int maxDepth = Math.Max(lastRow, lastColumn);

            int foundRowNum = -1;
            int foundColumnNum = -1;

            for (int depthNum = 1; depthNum <= maxDepth; ++depthNum)
            {
                if (-1 == foundRowNum && 0 < worksheet.Application.WorksheetFunction.CountA(worksheet.Rows[depthNum]))
                {
                    foundRowNum = depthNum;
                }

                if (-1 == foundColumnNum && 0 < worksheet.Application.WorksheetFunction.CountA(worksheet.Columns[depthNum]))
                {
                    foundColumnNum = depthNum;
                }

                if (-1 != foundRowNum && -1 != foundColumnNum)
                {
                    firstCell = worksheet.Cells[foundRowNum, foundColumnNum];
                    break;
                }
            }

            return firstCell;
        }

        private Range GetLastDataCell(Worksheet worksheet, Range firstDataCell_)
        {
            Range firstCell = worksheet.Cells[1, 1];
            int lastRow = firstCell.End[XlDirection.xlDown].Row;
            int lastColumn = firstCell.End[XlDirection.xlToRight].Column;

            Range lastCell = worksheet.Cells[lastRow, lastColumn];

            if (0 < worksheet.Application.WorksheetFunction.CountA(lastCell) || null == firstDataCell_)
            {
                return lastCell;
            }

            int foundRowNum = -1;
            int foundColumnNum = -1;

            for (int rowNum = firstDataCell_.Row + 1; rowNum < lastRow && -1 == foundRowNum; ++rowNum)
            {
                if (0 == worksheet.Application.WorksheetFunction.CountA(worksheet.Rows[rowNum]))
                {
                    foundRowNum = rowNum;
                }
            }

            for (int columnNum = firstDataCell_.Column + 1; columnNum < lastColumn && -1 == foundColumnNum; ++columnNum)
            {
                if (0 == worksheet.Application.WorksheetFunction.CountA(worksheet.Columns[columnNum]))
                {
                    foundColumnNum = columnNum;
                }
            }

            if (1 < foundRowNum && 1 < foundColumnNum)
            {
                lastCell = worksheet.Cells[foundRowNum - 1, foundColumnNum - 1];
            }

            return lastCell;
        }
    }
}
