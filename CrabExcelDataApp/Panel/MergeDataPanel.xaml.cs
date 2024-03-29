﻿using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service;
using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Store;
using CrabExcelDataApp.Validator;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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
        private readonly ExcelReader excelReader;
        private readonly ExcelWriter excelWriter;
        private readonly TableStore templateTableStore = TableStore.GetInstance();
        private string[] chosenFilePaths;
        private readonly ExcelFilterModel mergeFilterModel = new ExcelFilterModel()
        {
            isStandardTemplate = true,
            isFilterIgnoreHiddenRows = false,
        };
        private string fileToSavePath;

        public MergeDataPanel()
        {
            InitializeComponent();

            /// LOGGER
            logHelper = new LogHelper(this);
            logHelper.SetLogListView(listViewLog);
            logHelper.Debug(">> Start Merge App >>");

            excelReader = new ExcelReader(logHelper);
            excelWriter = new ExcelWriter(logHelper);

            btnTemplateFile.Click += BtnTemplateFile_Click;
            btnPartialFiles.Click += BtnPartialFiles_Click;
            btnMerge.Click += BtnMerge_Click;
        }

        private void BtnTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load File Excel >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog()
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
                    List<TableModel> readData = excelReader.ReadData<string>(chosenFilePath);
                    templateTableStore.SetData(readData);

                    UpdateDataForUI();
                }
            }
        }

        private void BtnPartialFiles_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load Partial File Excel >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Choose a Excel File",
                Multiselect = true,
                RestoreDirectory = true,
            };

            if (true == openFileDialog.ShowDialog())
            {
                chosenFilePaths = openFileDialog.FileNames;
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

            var headers = templateTableStore.GetSheetAt(0).GetHeader();

            if (null == headers || 0 == headers.Count || 0 == headers.Count)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Template headers are empty",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
                return;
            }


            this.mergeFilterModel.isStandardTemplate = excelFilter__chxStandardTemplate.IsChecked ?? false;
            this.mergeFilterModel.isFilterIgnoreHiddenRows = excelFilter__chxIgnoreHiddenRows.IsChecked ?? false;

            logHelper.Info($"Do Merge Data for [{string.Join(",", chosenFilePaths)}]");

            var fileToSavePath = "";
            do
            {
                fileToSavePath = ChooseFileToSave();

                if (string.IsNullOrEmpty(fileToSavePath))
                {
                    var dialogResult = System.Windows.Forms.MessageBox.Show(
                        "Please choose a file to save",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OKCancel,
                        System.Windows.Forms.MessageBoxIcon.Information
                    );

                    if (System.Windows.Forms.DialogResult.Cancel == dialogResult)
                    {
                        break;
                    }
                }
            } while (string.IsNullOrEmpty(fileToSavePath));


            if (!string.IsNullOrEmpty(fileToSavePath))
            {
                StartBackgroundWorker(fileToSavePath);
            }
            else
            {
                logHelper.Warn("Stop execution because you dont choose saved file.");
            }
        }

        private void StartBackgroundWorker(string fileToSavePath)
        {
            this.fileToSavePath = fileToSavePath;
            btnMerge.IsEnabled = false;
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);

            backgroundWorker.RunWorkerAsync(new MergeBackgroundModel
            {
                chosenPartialFilePaths = chosenFilePaths,
                templateTableStore = templateTableStore,
                fileToSavePath = fileToSavePath
            });
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            logHelper.Info($"Progress percent: {e.ProgressPercentage}%");
            processBar.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnMerge.IsEnabled = true;
            processBar.Value = 100;
            logHelper.Info("Progress percent: 100% - Completed");

            var dialogResult = System.Windows.Forms.MessageBox.Show(
                "Completed. Do you want to open containing folder of saved file?",
                "Information",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Information
            );

            if (System.Windows.Forms.DialogResult.Yes == dialogResult)
            {
                Process.Start("explorer.exe", "/select," + fileToSavePath);
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker backgroundWorker = sender as BackgroundWorker;
            DoMergeData(backgroundWorker, e.Argument as MergeBackgroundModel);
        }

        private void DoMergeData(BackgroundWorker backgroundWorker, MergeBackgroundModel mergeBackgroundModel)
        {
            backgroundWorker.ReportProgress(0);
            int totalFileCount = mergeBackgroundModel.chosenPartialFilePaths.Length;

            var templateHeader = mergeBackgroundModel.templateTableStore.GetSheetAt(0).GetHeader();
            MergeDataService mergeDataService = new MergeDataService(this.logHelper, this.excelReader);
            for (int fileIdx = 0; fileIdx < totalFileCount; ++fileIdx)
            {
                mergeDataService.AddPartialDataFile(
                    templateHeader,
                    mergeBackgroundModel.chosenPartialFilePaths[fileIdx],
                    this.mergeFilterModel
                );

                float workPercent = fileIdx * 100 / totalFileCount;

                backgroundWorker.ReportProgress((int)workPercent);
            }

            var totalData = mergeDataService.TotalData;

            if (!StringValidator.IsBlank(mergeBackgroundModel.fileToSavePath))
            {
                excelWriter.WriteToFile(mergeBackgroundModel.fileToSavePath, "total_data", mergeBackgroundModel.templateTableStore.GetSheetAt(0).GetHeader(), totalData);
            }
        }

        private string ChooseFileToSave()
        {
            logHelper.Debug(">> Start Save Total File Excel >>");

            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Save Output Excel File",
                RestoreDirectory = true,
            };

            if (true == saveFileDialog.ShowDialog())
            {
                return saveFileDialog.FileName;
            }
            else
            {
                return string.Empty;
            }
        }
    }
}
