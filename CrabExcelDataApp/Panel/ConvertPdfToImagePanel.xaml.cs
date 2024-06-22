using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service;
using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Store;
using CrabExcelDataApp.Validator;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace CrabExcelDataApp.Panel
{
    /// <summary>
    /// Interaction logic for MergeDataPanel.xaml
    /// </summary>
    public partial class ConvertPdfToImagePanel : UserControl
    {
        private readonly LogHelper logHelper;
        private string imageOutputFolderPath;

        public ConvertPdfToImagePanel()
        {
            InitializeComponent();

            /// LOGGER
            logHelper = new LogHelper(this);
            logHelper.SetLogListView(listViewLog);
            logHelper.Debug(">> Start Convert Pdf to Images App >>");

            btnPdfFile.Click += BtnPdfFileBrowse_Click;
            btnOutputFolder.Click += BtnOutputFolder_Click;
            btnConvert.Click += BtnConvert_Click;
        }

        private void BtnOutputFolder_Click(object sender, RoutedEventArgs e)
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
                inpOutputFolder.Text = openFolderDialog.SelectedPath;
            }
        }

        private void BtnPdfFileBrowse_Click(object sender, RoutedEventArgs e)
        {
            logHelper.Debug(">> Start Load PDF File >>");

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "Choose a PDF File",
                RestoreDirectory = true
            };

            if (true == openFileDialog.ShowDialog())
            {
                string chosenFilePath = openFileDialog.FileName;
                logHelper.Debug($"<< Chosen File: {chosenFilePath} <<");
                inpPdfFile.Text = chosenFilePath;
            }
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            string folderToSavePath = inpOutputFolder.Text;
            if (StringValidator.IsBlank(folderToSavePath))
            {
                System.Windows.MessageBox.Show("Please select a folder to save", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string pdfFilePath = inpPdfFile.Text;
            if (StringValidator.IsBlank(pdfFilePath))
            {
                System.Windows.MessageBox.Show("Please select PDF file to convert", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }


            StartBackgroundWorker(pdfFilePath, folderToSavePath);
        }

        private void StartBackgroundWorker(string pdfFilePath, string saveFolderPath)
        {
            btnConvert.IsEnabled = false;
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);

            this.imageOutputFolderPath = saveFolderPath;
            backgroundWorker.RunWorkerAsync(new ConvertPdfToImageBackgroundModel
            {
                pdfFilePath = pdfFilePath,
                imageOutputFolderPath = saveFolderPath,
            });
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            logHelper.Info($"Progress percent: {e.ProgressPercentage}%");
            processBar.Value = e.ProgressPercentage;
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnConvert.IsEnabled = true;
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
                Process.Start("explorer.exe", "/select," + this.imageOutputFolderPath);
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker backgroundWorker = sender as BackgroundWorker;
            DoConvertPdfToImage(backgroundWorker, e.Argument as ConvertPdfToImageBackgroundModel);
        }

        private void DoConvertPdfToImage(BackgroundWorker bgw, ConvertPdfToImageBackgroundModel model_)
        {
            logHelper.Info($"Convert PDF [{model_.pdfFilePath}] to folder: {model_.imageOutputFolderPath}");

            bgw.ReportProgress(0);

            using (var document = PdfDocument.Load(model_.pdfFilePath))
            {
                Directory.CreateDirectory(model_.imageOutputFolderPath);

                int pageCount = document.PageCount;

                for (int pageIdx = 0; pageIdx < pageCount; pageIdx++)
                {
                    using (var image = document.Render(pageIdx, 300, 300, true))
                    {
                        image.Save(Path.Combine(model_.imageOutputFolderPath, $"page_{pageIdx}.png"), ImageFormat.Png);
                    }
                    var percent_ = (float)(pageIdx + 1) * 100 / pageCount;
                    bgw.ReportProgress((int)percent_);
                }
            }
        }
    }
}
