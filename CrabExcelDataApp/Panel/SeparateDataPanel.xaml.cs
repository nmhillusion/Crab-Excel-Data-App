using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service;
using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Store;
using CrabExcelDataApp.Validator;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace CrabExcelDataApp.Panel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class SeparateDataPanel : UserControl
    {
        private readonly LogHelper logHelper;
        private readonly ExcelReader excelReader;
        private readonly ExcelWriter excelWriter;
        private readonly TableStore tableStore = TableStore.GetInstance();
        private readonly DifferenceService differenceService = new DifferenceService();

        public SeparateDataPanel()
        {
            InitializeComponent();

            /// LOGGER
            logHelper = new LogHelper(this);
            logHelper.SetLogListView(listViewLog);
            logHelper.Debug(">> Start Separate App >>");

            excelReader = new ExcelReader(logHelper);
            excelWriter = new ExcelWriter(logHelper);

            /// EVENT
            btnFileToSeparate.Click += BtnSelectFile_Click;
            cboSheetIdx.SelectionChanged += CboSheetIdx_SelectionChanged;
            btnFolderToSave.Click += BtnFolderToSave_Click;

            btnSeparate.Click += BtnSeparate_Click;
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load File Excel >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Choose a Excel File"
            };

            if (true == openFileDialog.ShowDialog())
            {
                string chosenFilePath = openFileDialog.FileName;
                logHelper.Debug($"<< Chosen File: {chosenFilePath} <<");
                inpFileToSeparate.Text = chosenFilePath;

                if (!StringValidator.IsBlank(chosenFilePath))
                {
                    List<TableModel> readData = excelReader.ReadData<string>(chosenFilePath);
                    tableStore.SetData(readData);

                    UpdateDataForUI();
                }
            }
        }

        private void BtnFolderToSave_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Chosing a Folder to save >>");

            System.Windows.Forms.FolderBrowserDialog openFolderDialog = new System.Windows.Forms.FolderBrowserDialog()
            {
                ShowNewFolderButton = true,
                Description = "Choose a Folder to Save"
            };

            _ = openFolderDialog.ShowDialog();
            if (!StringValidator.IsBlank(openFolderDialog.SelectedPath))
            {
                inpFolderToSave.Text = openFolderDialog.SelectedPath;
            }
        }

        private void BtnSeparate_Click(object sender, RoutedEventArgs e)
        {
            int selectedSheetIdx = cboSheetIdx.SelectedIndex;
            int selectedColumnIdx = cboColumnIdx.SelectedIndex;
            string folderToSavePath = inpFolderToSave.Text;

            if (0 == tableStore.Count)
            {
                System.Windows.MessageBox.Show("Please select a file to separate", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (-1 == selectedSheetIdx)
            {
                System.Windows.MessageBox.Show("Please select a sheet to separate", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (-1 == selectedColumnIdx)
            {
                System.Windows.MessageBox.Show("Please select a column to separate", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (StringValidator.IsBlank(folderToSavePath))
            {
                System.Windows.MessageBox.Show("Please select a folder to save", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            StartBackgroundWorker();
        }

        private void StartBackgroundWorker()
        {
            btnSeparate.IsEnabled = false;
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);

            SeparateBackgroundModel separateBackgroundModel = new SeparateBackgroundModel()
            {
                selectedSheetIdx = cboSheetIdx.SelectedIndex,
                selectedColumnIdx = cboColumnIdx.SelectedIndex,
                folderToSavePath = inpFolderToSave.Text
            };
            backgroundWorker.RunWorkerAsync(separateBackgroundModel);
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            logHelper.Info($"Progress percent: {e.ProgressPercentage}%");
            processBar.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSeparate.IsEnabled = true;
            processBar.Value = 100;
            logHelper.Info("Progress percent: 100% - Completed");

            System.Windows.Forms.MessageBox.Show("Completed", "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker backgroundWorker = sender as BackgroundWorker;
            DoSeparateData(backgroundWorker, e.Argument as SeparateBackgroundModel);
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

                List<object> headers = sheetToPush.GetHeader();
                if (0 < headers.Count)
                {
                    for (int headerCellIdx = 0; headerCellIdx < headers.Count; ++headerCellIdx)
                    {
                        cboColumnIdx.Items.Add(
                            $"{headerCellIdx} - {headers[headerCellIdx]}"
                        );
                    }
                    if (0 < headers.Count)
                    {
                        cboColumnIdx.SelectedIndex = 0;
                    }
                }
            }
        }

        private void DoSeparateData(BackgroundWorker backgroundWorker, SeparateBackgroundModel separateBackgroundModel)
        {
            int selectedSheetIdx = separateBackgroundModel.selectedSheetIdx;
            int selectedColumnIdx = separateBackgroundModel.selectedColumnIdx;
            string folderToSavePath = separateBackgroundModel.folderToSavePath;

            TableModel sheetToSeparate = tableStore.GetSheetAt(selectedSheetIdx);
            Debug.WriteLine($"sheetToSeparate : {string.Join(", ", sheetToSeparate.GetBody())}");
            List<object> dataAtColumnIdx = sheetToSeparate.GetDataAtColumnIdx(selectedColumnIdx);
            Debug.WriteLine($"dataAtColumnIdx : {string.Join(", ", dataAtColumnIdx)}");

            List<string> allDistinctDataOfSeparateData = new List<string>();
            allDistinctDataOfSeparateData.AddRange(
                differenceService.DistinctListObject(
                    dataAtColumnIdx
                )
            );

            logHelper.Info("----------- allDistinctDataOfSeparateData ---------------- : " + string.Join(", ", allDistinctDataOfSeparateData));
            for (int diffItemIdx = 0; diffItemIdx < allDistinctDataOfSeparateData.Count; ++diffItemIdx)
            {
                backgroundWorker.ReportProgress(diffItemIdx * 100 / allDistinctDataOfSeparateData.Count);

                string diffItem = allDistinctDataOfSeparateData[diffItemIdx];
                logHelper.Info("diffItem: " + diffItem);

                /// CREATE FILE FOR EACH DIFF ITEM
                string pathToSaveFile = Path.Combine(folderToSavePath, diffItem + ".xlsx");
                if (File.Exists(pathToSaveFile))
                {
                    File.Delete(pathToSaveFile);
                }

                /// FILTER DATA FOR THIS DIFF ITEM
                List<List<object>> filteredData = differenceService.FilterData(sheetToSeparate, selectedColumnIdx, diffItem);

                /// SAVE FILTERED DATA TO EXCEL FILE
                excelWriter.WriteToFile(pathToSaveFile, diffItem, sheetToSeparate.GetHeader(), filteredData);
            }
        }
    }
}
