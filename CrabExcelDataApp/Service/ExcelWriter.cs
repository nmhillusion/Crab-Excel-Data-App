using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Util;
using System;
using System.Collections.Generic;
using System.Diagnostics;

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
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            try
            {
                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel.Worksheet oSheet = oWB.Sheets.Add();

                oSheet.Name = sheetName.Length >= 31 ? sheetName.Substring(0, 30) : sheetName;

                SaveDataToCells(oSheet.Cells, headers, bodyData);

                oWB.SaveAs(excelPathToSave);

                logHelper.Info($"Export done for sheet {sheetName} to path: {excelPathToSave}!");
            }
            catch (Exception ex)
            {
                logHelper.Error(ex);
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close();
                }
            }
        }

        private void SaveDataToCells(Microsoft.Office.Interop.Excel.Range rangeExcelToSave, List<object> headers, List<List<object>> bodyData)
        {
            int row = 1;
            int col = 1;

            /// SAVE HEADERs
            {
                foreach (object iHeader in headers)
                {
                    rangeExcelToSave[row, col] = iHeader;
                    col += 1;
                }
                col = 1;
                row += 1;
            }

            /// SAVE BODY DATA
            foreach (List<object> rowData in bodyData)
            {
                Debug.WriteLine(this.GetType().Name + " - read data from excel: " + string.Join(", ", rowData));
                foreach (object cellData in rowData)
                {
                    rangeExcelToSave[row, col].NumberFormat = "@";
                    rangeExcelToSave[row, col] = StringUtil.ToString(cellData);
                    col += 1;
                }
                col = 1;
                row += 1;
            }
        }
    }
}
