using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Util;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CrabExcelDataApp.Service
{
    class MergeDataService
    {
        private readonly ExcelReader reader = new ExcelReader();
        private readonly LogHelper logHelper;

        private List<List<object>> Data;

        public List<List<object>> TotalData
        {
            get
            {
                var cloneData = new List<List<object>>();
                cloneData.AddRange(Data);
                return cloneData;
            }
        }

        public MergeDataService(LogHelper logHelper)
        {
            Data = new List<List<object>>();
            this.logHelper = logHelper;
        }

        private bool IsValidPartialData(List<string> templateHeader, Model.TableModel sheet_)
        {
            bool result = true;
            var sheetHeaders = sheet_.GetHeader();

            foreach (var headerColumn in templateHeader)
            {
                string headerColumn_ = headerColumn.ToString().Trim();
                if (!sheetHeaders.Any(it => headerColumn_.Equals(it.ToString().Trim())))
                {
                    result = false;
                }
            }

            return result;
        }

        private Dictionary<string, int> MappingHeadersWithColumns(List<string> templateHeader, Model.TableModel sheet_)
        {
            Dictionary<string, int> resultDict = new Dictionary<string, int>();
            var sheetHeaders = sheet_.GetHeader();

            foreach (var headerColumn in templateHeader)
            {
                string headerColumn_ = headerColumn.ToString().Trim();
                int idx = sheetHeaders.FindIndex(it => headerColumn_.Equals(it.ToString().Trim()));

                if (-1 == idx)
                {
                    throw new Exception($"Cannot find columnIdx of header: {headerColumn} in sheet: {sheet_.tableName}");
                }

                resultDict[headerColumn_] = idx;
            }

            return resultDict;
        }

        public MergeDataService AddPartialDataFile(List<object> templateHeader, string partialFilePath)
        {
            List<Model.TableModel> sheets = reader.ReadData(partialFilePath);
            List<string> templateHeader_ = templateHeader.Select(it => it.ToString().Trim()).ToList();

            foreach (var sheet_ in sheets)
            {
                if (IsValidPartialData(templateHeader_, sheet_))
                {
                    Dictionary<string, int> mappingHeadersColumns = MappingHeadersWithColumns(templateHeader_, sheet_);

                    var sheetData = sheet_.GetBody();
                    foreach (var row_ in sheetData)
                    {
                        List<object> newRowData = new List<object>();

                        foreach (var headerName in templateHeader_)
                        {
                            var colIdx = mappingHeadersColumns[headerName];
                            newRowData.Add(StringUtil.ToString(row_[colIdx]));
                        }

                        Data.Add(newRowData);
                    }
                }
                else
                {
                    logHelper.Error($"Sheet is not valid {sheet_.tableName} - file: {partialFilePath}");
                }
            }

            return this;
        }
    }
}
