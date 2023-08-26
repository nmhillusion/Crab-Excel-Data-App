using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service.Logger;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Controls;

namespace CrabExcelDataApp.Service
{
    class ExcelReader
    {
        private readonly LogHelper logHelper;

        public ExcelReader() : this(null)
        {
        }

        public ExcelReader(LogHelper logHelper)
        {
            if (null != logHelper)
            {
                this.logHelper = logHelper;
            }
            else
            {
                this.logHelper = new LogHelper(this);
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public List<TableModel> ReadData<T>(string excelFilePath)
        {
            return ReadData<T>(excelFilePath, new ExcelFilterModel()
            {
                isStandardTemplate = true,
                isFilterIgnoreHiddenRows = false,
            });
        }

        public List<TableModel> ReadData<T>(string excelFilePath, ExcelFilterModel excelFilterModel)
        {
            return ReadData<T>(excelFilePath, excelFilterModel, null);
        }

        public List<TableModel> ReadData<T>(string excelFilePath, ExcelFilterModel excelFilterModel, List<T> startHeaderTemplate)
        {
            if (excelFilterModel.isStandardTemplate && !excelFilterModel.isFilterIgnoreHiddenRows)
            {
                return ReadDataAllRows(excelFilePath, excelFilterModel, startHeaderTemplate);
            }
            else
            {
                return ReadDataWithAdvantageFilters(excelFilePath, excelFilterModel, startHeaderTemplate);
            }
        }

        private int FindStartRowNum<T>(System.Data.DataTable sheetData, List<T> startHeaderTemplate)
        {
            int rowNum = -1;

            if (null == startHeaderTemplate || 0 == startHeaderTemplate.Count)
            {
                return 0;
            }

            var rows = sheetData.Rows;
            for (int rowIdx = 0; rowIdx < rows.Count; ++rowIdx)
            {
                var rowData = rows[rowIdx];
                object[] cells = rowData.ItemArray;
                List<string> tmpList = new List<object>(cells).Select(it => it.ToString().Trim()).ToList();
                List<object> list_ = new List<object>(tmpList);

                if (startHeaderTemplate.All(it => list_.Contains(it)))
                {
                    rowNum = rowIdx + 1;
                    break;
                }
            }

            logHelper.Info($"Found started row num is: {rowNum}");
            return rowNum;
        }

        private int FindStartRowNum<T>(Worksheet sheetData, List<T> startHeaderTemplate)
        {
            int rowNum = -1;

            if (null == startHeaderTemplate || 0 == startHeaderTemplate.Count)
            {
                return 0;
            }

            int rowCount = sheetData.UsedRange.Rows.Count;
            for (int rowNum_ = sheetData.UsedRange.Row; rowNum_ <= rowCount; ++rowNum_)
            {
                //logHelper.Debug($"Checking on row: {rowNum_}");
                Range rowRange = sheetData.Rows[rowNum_];
                object[,] rowData2 = rowRange.Value2;
                List<object> firstRowData = new List<object>();
                for (int cellCol = 1; cellCol <= rowData2.Length; ++cellCol)
                {
                    var cellData = rowData2[1, cellCol];

                    if (null != cellData && !string.IsNullOrEmpty(cellData.ToString()))
                    {
                        firstRowData.Add(cellData.ToString().Trim());
                    }
                }

                if (startHeaderTemplate.All(it => firstRowData.Contains(it)))
                {
                    rowNum = rowNum_;
                    break;
                }
            }

            logHelper.Info($"Found started row num is: {rowNum}");
            return rowNum;
        }

        public List<TableModel> ReadDataAllRows<T>(string excelFilePath, ExcelFilterModel excelFilterModel, List<T> startHeaderTemplate)
        {
            logHelper.Info("Read Excel at " + excelFilePath);
            string excelFileName = Path.GetFileName(excelFilePath);

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

                    int startRowNum = FindStartRowNum(dataTable, startHeaderTemplate);
                    if (-1 == startRowNum)
                    {
                        logHelper.Warn("This sheet is not valid, jump to the next sheet.");
                        continue;
                    }


                    DataRowCollection dataRowCollection = dataTable.Rows;
                    int startRowIndex = Math.Max(0, startRowNum - 1);

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
                        logHelper.Info($"[{excelFileName}][{dataTable.TableName}][{rowIdx + 1}] read data from excel: " + string.Join(", ", rowData));

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

        public List<TableModel> ReadDataWithAdvantageFilters<T>(string excelFilePath, ExcelFilterModel excelFilterModel, List<T> startHeaderTemplate)
        {
            logHelper.Info("Read Excel at " + excelFilePath);
            string excelFileName = Path.GetFileName(excelFilePath);

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
                    int startRowNum_ = FindStartRowNum(worksheet, startHeaderTemplate);
                    if (-1 == startRowNum_)
                    {
                        logHelper.Warn("This sheet is not valid, jump to the next sheet.");
                        continue;
                    }

                    Range firstCell_ = GetFirstDataCell(worksheet, excelFilterModel, startRowNum_);
                    Range lastCell_ = GetLastDataCell(worksheet, excelFilterModel, firstCell_, startRowNum_);
                    logHelper.Info($"first cell: [{firstCell_.Row}, {firstCell_.Column}]");
                    logHelper.Info($"last cell: [{lastCell_.Row}, {lastCell_.Column}]");

                    logHelper.Info($"lastRow: {lastCell_.Row}");
                    logHelper.Info($"lastColumn: {lastCell_.Column}");

                    List<List<object>> sheetData = new List<List<object>>();
                    int startRowNum = Math.Max(firstCell_.Row, startRowNum_);

                    for (int rowNum = startRowNum; rowNum <= lastCell_.Row; ++rowNum)
                    {
                        Range row_ = worksheet.Rows[rowNum];

                        if (IsValidHiddenCondition(row_, excelFilterModel))
                        {
                            List<object> rowData = new List<object>();
                            for (int colNum = firstCell_.Column; colNum <= lastCell_.Column; ++colNum)
                            {
                                Range cell_ = worksheet.Cells[rowNum, colNum];
                                rowData.Add(cell_.Value2);
                            }
                            logHelper.Info($"[{excelFileName}][{worksheet.Name}][{rowNum}] read data from excel: " + string.Join(", ", rowData));
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

        private Range GetFirstDataCell(Worksheet worksheet, ExcelFilterModel excelFilterModel, int startRowNum)
        {
            int startRow = Math.Max(1, startRowNum);
            Range firstCell = worksheet.Cells[startRow, 1];

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
                int currentCheckingRowNum = Math.Max(depthNum, startRow);
                if (-1 == foundRowNum && 0 < worksheet.Application.WorksheetFunction.CountA(worksheet.Rows[currentCheckingRowNum]))
                {
                    foundRowNum = depthNum;
                }

                int currentCheckingColumnNum = depthNum;
                if (-1 == foundColumnNum && 0 < worksheet.Application.WorksheetFunction.CountA(worksheet.Columns[currentCheckingColumnNum]))
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

        private Range GetLastDataCell(Worksheet worksheet, ExcelFilterModel excelFilterModel, Range firstDataCell_, int startRowNum)
        {
            int startRow = Math.Max(1, startRowNum);
            Range firstCell = worksheet.Cells[startRow, 1];
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
