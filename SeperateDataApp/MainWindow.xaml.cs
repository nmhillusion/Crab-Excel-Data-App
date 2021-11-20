using Microsoft.Win32;
using SeperateDataApp.Service;
using SeperateDataApp.Store;
using SeperateDataApp.Validator;
using System.Collections.Generic;
using System.Windows;

namespace SeperateDataApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly LogHelper logHelper;
        private OpenFileDialog openFileDialog;
        private readonly ExcelReader excelReader = new();

        public MainWindow()
        {
            logHelper = new LogHelper(this);

            logHelper.Debug(">> Start App >>");
            InitializeComponent();
            btnSelectFile.Click += BtnSelectFile_Click;
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load File Excel >>");

            openFileDialog = new OpenFileDialog()
            {
                Filter = "Excel 2007 or newer (*.xlsx)|*.xlsx|Prior of Excel 2007 (*.xls)|*.xls",
                Title = "Choose a Excel File"
            };

            if (true == openFileDialog.ShowDialog())
            {
                string chosenFilePath = openFileDialog.FileName;
                logHelper.Debug($"<< Chosen File: { chosenFilePath } <<");

                if (!StringValidator.IsBlank(chosenFilePath))
                {
                    List<List<List<string>>> readData = excelReader.ReadData(chosenFilePath);
                    TableStore.GetInstance().SetData(readData);

                    List<List<string>> firstSheet = TableStore.GetInstance().GetSheetAt(0);
                    List<string> header = firstSheet[0];
                    foreach (string headerCell in header)
                    {
                        logHelper.Debug($"header -> {headerCell}");
                    }
                }
            }
        }
    }
}
