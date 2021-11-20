using SeperateDataApp.Service;
using SeperateDataApp.Store;
using SeperateDataApp.Validator;
using System.Collections.Generic;
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
        private readonly TableStore tableStore = TableStore.GetInstance();
        private readonly DifferenceService differenceService = new();

        public MainWindow()
        {
            logHelper = new LogHelper(this);

            logHelper.Debug(">> Start App >>");
            InitializeComponent();
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

            if (0 == tableStore.Count)
            {
                System.Windows.MessageBox.Show("Error", "Please select a file to seperate", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (-1 == selectedSheetIdx)
            {
                System.Windows.MessageBox.Show("Error", "Please select a sheet to seperate", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (-1 == selectedColumnIdx)
            {
                System.Windows.MessageBox.Show("Error", "Please select a column to seperate", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DoSeperateData();
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

                List<object> header = sheetToPush.GetHeader();
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

        private void DoSeperateData()
        {
            int selectedSheetIdx = cboSheetIdx.SelectedIndex;
            int selectedColumnIdx = cboColumnIdx.SelectedIndex;

            TableModel sheetToSeperate = tableStore.GetSheetAt(selectedSheetIdx);
            List<object> dataAtColumnIdx = sheetToSeperate.GetDataAtColumnIdx(selectedColumnIdx);
            ISet<string> allDistinctDataOfSeperateData = differenceService.distinctListObject(
                dataAtColumnIdx
            );

            logHelper.Info("----------- allDistinctDataOfSeperateData ---------------- " + allDistinctDataOfSeperateData.Count);
            foreach (string diffItem in allDistinctDataOfSeperateData)
            {
                logHelper.Info("\t diffItem: " + diffItem);
            }
        }
    }
}
