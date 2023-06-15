using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service;
using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Store;
using CrabExcelDataApp.Validator;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace CrabExcelDataApp.Panel
{
    /// <summary>
    /// Interaction logic for MergeDataPanel.xaml
    /// </summary>
    public partial class MergeDataPanel : UserControl
    {
        private readonly LogHelper logHelper;
        private readonly ExcelReader excelReader = new();
        private readonly ExcelWriter excelWriter = new();
        private readonly TableStore templateTableStore = TableStore.GetInstance();
        private string[] chosenFilePaths;

        public MergeDataPanel()
        {
            InitializeComponent();

            /// LOGGER
            logHelper = new LogHelper(this);
            logHelper.SetLogListView(listViewLog);
            logHelper.Debug(">> Start Merge App >>");

            btnTemplateFile.Click += BtnTemplateFile_Click;
            btnPartialFiles.Click += BtnPartialFiles_Click;
            btnMerge.Click += BtnMerge_Click;
        }

        private void BtnTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load File Excel >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Choose a Excel File",
                RestoreDirectory = true
            };

            if (true == openFileDialog.ShowDialog())
            {
                string chosenFilePath = openFileDialog.FileName;
                logHelper.Debug($"<< Chosen File: {chosenFilePath} <<");
                inpTemplateFile.Text = chosenFilePath;

                if (!StringValidator.IsBlank(chosenFilePath))
                {
                    List<TableModel> readData = excelReader.ReadData(chosenFilePath);
                    templateTableStore.SetData(readData);

                    UpdateDataForUI();
                }
            }
        }

        private void BtnPartialFiles_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load Partial File Excel >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Choose a Excel File",
                Multiselect = true,
                RestoreDirectory = true,
            };

            if (true == openFileDialog.ShowDialog())
            {
                chosenFilePaths = openFileDialog.SafeFileNames;
                logHelper.Debug($"<< Chosen partial files: {string.Join(";", chosenFilePaths)} <<");
                inpPartialFiles.Text = string.Join(";", chosenFilePaths);
            }
        }

        private void UpdateDataForUI()
        {
            logHelper.Info("Update GUI : tableStore count - " + templateTableStore.Count);
        }

        private void BtnMerge_Click(object sender, RoutedEventArgs e)
        {
            if (0 == templateTableStore.Count)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Template is empty",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
                return;
            }

            if (null == chosenFilePaths || 0 == chosenFilePaths.Length)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Partial files are empty",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
                return;
            }

            DoMergeData();
        }

        private void DoMergeData()
        {
            logHelper.Info($"Do Merge Data for [{string.Join(",", chosenFilePaths)}]");
        }
    }
}
