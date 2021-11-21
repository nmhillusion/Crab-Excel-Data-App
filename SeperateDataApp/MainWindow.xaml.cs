using Microsoft.Toolkit.Uwp.Notifications;
using SeperateDataApp.Model;
using SeperateDataApp.Service;
using SeperateDataApp.Service.Logger;
using SeperateDataApp.Store;
using SeperateDataApp.Validator;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Forms;

namespace SeperateDataApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly LogHelper logHelper;
        private readonly ExcelReader excelReader = new();
        private readonly ExcelWriter excelWriter = new();
        private readonly TableStore tableStore = TableStore.GetInstance();
        private readonly DifferenceService differenceService = new();

        public MainWindow()
        {
            InitializeComponent();

            /// LOGGER
            logHelper = new LogHelper(this);
            logHelper.SetLogListView(listViewLog);
            logHelper.Debug(">> Start App >>");

            /// EVENT
            btnFileToSeperate.Click += BtnSelectFile_Click;
            cboSheetIdx.SelectionChanged += CboSheetIdx_SelectionChanged;
            btnFolderToSave.Click += BtnFolderToSave_Click;

            btnSeperate.Click += BtnSeperate_Click;
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load File Excel >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Choose a Excel File"
            };

            if (true == openFileDialog.ShowDialog())
            {
                string chosenFilePath = openFileDialog.FileName;
                logHelper.Debug($"<< Chosen File: { chosenFilePath } <<");
                inpFileToSeperate.Text = chosenFilePath;

                if (!StringValidator.IsBlank(chosenFilePath))
                {
                    List<TableModel> readData = excelReader.ReadData(chosenFilePath);
                    tableStore.SetData(readData);

                    UpdateDataForUI();
                }
            }
        }

        private void BtnFolderToSave_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Chosing a Folder to save >>");

            FolderBrowserDialog openFolderDialog = new()
            {
                ShowNewFolderButton = true,
                Description = "Choose a Folder to Save",
                UseDescriptionForTitle = true
            };

            _ = openFolderDialog.ShowDialog();
            if (!StringValidator.IsBlank(openFolderDialog.SelectedPath))
            {
                inpFolderToSave.Text = openFolderDialog.SelectedPath;
            }
        }

        private void BtnSeperate_Click(object sender, RoutedEventArgs e)
        {
            int selectedSheetIdx = cboSheetIdx.SelectedIndex;
            int selectedColumnIdx = cboColumnIdx.SelectedIndex;
            string folderToSavePath = inpFolderToSave.Text;

            if (0 == tableStore.Count)
            {
                new ToastContentBuilder()
                    .AddText("Error", AdaptiveTextStyle.Title)
                    .AddText("Please select a file to seperate", AdaptiveTextStyle.Body)
                    .Show();
                return;
            }

            if (-1 == selectedSheetIdx)
            {
                new ToastContentBuilder()
                    .AddText("Error", AdaptiveTextStyle.Title)
                    .AddText("Please select a sheet to seperate", AdaptiveTextStyle.Body)
                    .Show();
                return;
            }

            if (-1 == selectedColumnIdx)
            {
                new ToastContentBuilder()
                    .AddText("Error", AdaptiveTextStyle.Title)
                    .AddText("Please select a column to seperate", AdaptiveTextStyle.Body)
                    .Show();
                return;
            }

            if (StringValidator.IsBlank(folderToSavePath))
            {
                new ToastContentBuilder()
                    .AddText("Error", AdaptiveTextStyle.Title)
                    .AddText("Please select a folder to save", AdaptiveTextStyle.Body)
                    .Show();
                return;
            }

            StartBackgroundWorker();
        }

        private void StartBackgroundWorker()
        {
            btnSeperate.IsEnabled = false;
            BackgroundWorker backgroundWorker = new();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);

            SeperateBackgroundModel seperateBackgroundModel = new()
            {
                selectedSheetIdx = cboSheetIdx.SelectedIndex,
                selectedColumnIdx = cboColumnIdx.SelectedIndex,
                folderToSavePath = inpFolderToSave.Text
            };
            backgroundWorker.RunWorkerAsync(seperateBackgroundModel);
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            logHelper.Info("Percent: " + e.ProgressPercentage);
            processBar.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSeperate.IsEnabled = true;
            processBar.Value = 100;

            new ToastContentBuilder()
                .AddText("Completed!")
                .Show();
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker backgroundWorker = sender as BackgroundWorker;
            DoSeperateData(backgroundWorker, e.Argument as SeperateBackgroundModel);
        }

        private void UpdateDataForUI()
        {
            for (int sheetIdx = 0; sheetIdx < tableStore.Count; ++sheetIdx)
            {
                TableModel sheet = tableStore.GetSheetAt(sheetIdx);
                cboSheetIdx.Items.Add(
                    $"{sheetIdx} - {sheet.tableName}"
                );
            }
            if (0 < tableStore.Count)
            {
                cboSheetIdx.SelectedIndex = 0;
            }
        }

        private void CboSheetIdx_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            UpdateColumnItemsComboBox();
        }

        private void UpdateColumnItemsComboBox()
        {
            int selectedSheetIdx = cboSheetIdx.SelectedIndex;
            if (0 <= selectedSheetIdx && 0 < tableStore.Count)
            {
                TableModel sheetToPush = tableStore.GetSheetAt(selectedSheetIdx);

                List<List<object>> headers = sheetToPush.GetHeader();
                if (0 < headers.Count)
                {
                    List<object> header = headers[0];

                    for (int headerCellIdx = 0; headerCellIdx < header.Count; ++headerCellIdx)
                    {
                        cboColumnIdx.Items.Add(
                            $"{headerCellIdx} - {header[headerCellIdx]}"
                        );
                    }
                    if (0 < header.Count)
                    {
                        cboColumnIdx.SelectedIndex = 0;
                    }
                }
            }
        }

        private void DoSeperateData(BackgroundWorker backgroundWorker, SeperateBackgroundModel seperateBackgroundModel)
        {
            int selectedSheetIdx = seperateBackgroundModel.selectedSheetIdx;
            int selectedColumnIdx = seperateBackgroundModel.selectedColumnIdx;
            string folderToSavePath = seperateBackgroundModel.folderToSavePath;

            TableModel sheetToSeperate = tableStore.GetSheetAt(selectedSheetIdx);
            List<object> dataAtColumnIdx = sheetToSeperate.GetDataAtColumnIdx(selectedColumnIdx);
            List<string> allDistinctDataOfSeperateData = new();
            allDistinctDataOfSeperateData.AddRange(
                differenceService.DistinctListObject(
                    dataAtColumnIdx
                )
            );

            logHelper.Info("----------- allDistinctDataOfSeperateData ---------------- " + allDistinctDataOfSeperateData.Count);
            for (int diffItemIdx = 0; diffItemIdx < allDistinctDataOfSeperateData.Count; ++diffItemIdx)
            {
                backgroundWorker.ReportProgress(diffItemIdx * 100 / allDistinctDataOfSeperateData.Count);

                string diffItem = allDistinctDataOfSeperateData[diffItemIdx];
                logHelper.Info("diffItem: " + diffItem);

                /// CREATE FILE FOR EACH DIFF ITEM
                string pathToSaveFile = Path.Combine(folderToSavePath, diffItem + ".xlsx");
                if (File.Exists(pathToSaveFile))
                {
                    File.Delete(pathToSaveFile);
                }

                /// FILTER DATA FOR THIS DIFF ITEM
                List<List<object>> filteredData = differenceService.FilterData(sheetToSeperate, selectedColumnIdx, diffItem);

                /// SAVE FILTERED DATA TO EXCEL FILE
                excelWriter.WriteToFile(pathToSaveFile, diffItem, sheetToSeperate.GetHeader(), filteredData);
            }
        }
    }
}
