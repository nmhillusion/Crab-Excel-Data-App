using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace CrabExcelDataApp.Service
{
    class ExcelWriter
    {
        private readonly LogHelper logHelper;

        public ExcelWriter()
        {
            logHelper = new LogHelper(this);
        }

        public void WriteToFile(string excelPathToSave, string sheetName, List<object> headers, List<List<object>> bodyData)
        {
            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel.Worksheet oSheet = null;
            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Add(Type.Missing);
                oSheet = oWB.Sheets.Add();

                oSheet.Name = sheetName.Length >= 31 ? sheetName.Substring(0, 30) : sheetName;

                SaveDataToCells(oSheet, headers, bodyData);

                oWB.SaveAs(excelPathToSave);

                logHelper.Info($"Export done for sheet {sheetName} to path: {excelPathToSave}!");
            }
            catch (Exception ex)
            {
                logHelper.Error(ex);
            }
            finally
            {
                oWB?.Close();
                oXL?.Quit();

                Marshal.ReleaseComObject(oSheet);
                Marshal.ReleaseComObject(oWB);
                Marshal.ReleaseComObject(oXL);
            }
        }

        private void SaveDataToCells(Microsoft.Office.Interop.Excel.Worksheet oSheet, List<object> headers, List<List<object>> bodyData)
        {
            int rowNum = 1;

            /// SAVE HEADERs
            {
                Microsoft.Office.Interop.Excel.Range startRange = oSheet.Cells[rowNum, 1];
                Microsoft.Office.Interop.Excel.Range endRange = oSheet.Cells[rowNum, headers.Count];
                oSheet.Range[startRange, endRange].Value = headers.ToArray();
                rowNum += 1;
            }

            /// SAVE BODY DATA
            {
                Microsoft.Office.Interop.Excel.Range startRange = oSheet.Cells[rowNum, 1];
                Microsoft.Office.Interop.Excel.Range endRange = oSheet.Cells[bodyData.Count + startRange.Row - 1, headers.Count];
                Microsoft.Office.Interop.Excel.Range dataRange_ = oSheet.Range[startRange, endRange];
                dataRange_.NumberFormat = "@";
                dataRange_.Value = ToArrayData(bodyData);
            }
        }

        private string[,] ToArrayData(List<List<object>> bodyData)
        {
            string[,] result_ = new string[bodyData.Count, bodyData.ElementAt(0).Count];
            for (int rowIdx = 0; rowIdx < bodyData.Count; ++rowIdx)
            {
                List<object> rowData = bodyData[rowIdx];
                for (int colIdx = 0; colIdx < rowData.Count; ++colIdx)
                {
                    result_[rowIdx, colIdx] = StringUtil.ToString(rowData[colIdx]);
                }
            }
            return result_;
        }
    }
}
