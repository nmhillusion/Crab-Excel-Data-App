using CrabExcelDataApp.Model;
using CrabExcelDataApp.Service;
using CrabExcelDataApp.Service.Logger;
using CrabExcelDataApp.Store;
using CrabExcelDataApp.Validator;
using PdfiumViewer;
using PdfSharp.Drawing;
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

            string dpi = inpDpi.Text;
            if (StringValidator.IsBlank(dpi) || 0 >= Int16.Parse(dpi))
            {
                System.Windows.MessageBox.Show("Please input valid DPI", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            ComboBoxItem selectedOutputTypeItem = (ComboBoxItem)cbxOutputType.SelectedItem;
            string selectedOutputTypeValue = selectedOutputTypeItem.Tag.ToString();

            StartBackgroundWorker(new ConvertPdfToImageBackgroundModel
            {
                pdfFilePath = pdfFilePath,
                imageOutputFolderPath = folderToSavePath,
                outputType = selectedOutputTypeValue,
                dpi = Int16.Parse(dpi),
            });
        }

        private void StartBackgroundWorker(ConvertPdfToImageBackgroundModel model_)
        {
            btnConvert.IsEnabled = false;
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);

            this.imageOutputFolderPath = model_.imageOutputFolderPath;
            backgroundWorker.RunWorkerAsync(model_);
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
            DoConvertPdfToAnotherType(backgroundWorker, e.Argument as ConvertPdfToImageBackgroundModel);
        }

        private void DoConvertPdfToAnotherType(BackgroundWorker bgw, ConvertPdfToImageBackgroundModel model_)
        {
            logHelper.Info($"Convert PDF [{model_.pdfFilePath}] to folder: {model_.imageOutputFolderPath} with outputType: {model_.outputType}");

            bgw.ReportProgress(0);

            if ("image".Equals(model_.outputType))
            {
                DoConvertPdfToImage(bgw, model_);
            }
            else
            {
                DoConvertPdfToPdf(bgw, model_);
            }
        }

        private void DoConvertPdfToImage(BackgroundWorker bgw, ConvertPdfToImageBackgroundModel model_)
        {
            using (var document = PdfDocument.Load(model_.pdfFilePath))
            {
                Directory.CreateDirectory(model_.imageOutputFolderPath);

                int pageCount = document.PageCount;

                for (int pageIdx = 0; pageIdx < pageCount; pageIdx++)
                {
                    using (var image = document.Render(pageIdx, model_.dpi, model_.dpi, true))
                    {
                        image.Save(Path.Combine(model_.imageOutputFolderPath, $"page_{pageIdx}.png"), ImageFormat.Png);
                    }
                    var percent_ = (float)(pageIdx + 1) * 100 / pageCount;
                    bgw.ReportProgress((int)percent_);
                }
            }
        }

        private void SavePagesToOnePdf(BackgroundWorker bgw, ConvertPdfToImageBackgroundModel model_, MemoryStream[] streamPages)
        {
            // Create a file stream for saving the PDF
            using (FileStream output = new FileStream(Path.Combine(model_.imageOutputFolderPath, "renderedDocument.pdf"), FileMode.Create))
            {
                using (PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument())
                {
                    int pageCount = streamPages.Length;
                    for (int pageIdx = 0; pageIdx < pageCount; ++pageIdx)
                    {
                        MemoryStream streamPage = streamPages[pageIdx];

                        using (streamPage)
                        {
                            byte[] byteData = streamPage.ToArray();

                            ///////////////////////
                            using (var image = XImage.FromStream(streamPage))
                            {
                                var page_ = document.AddPage();
                                page_.MediaBox = new PdfSharp.Pdf.PdfRectangle(new XRect(0, 0, image.PixelWidth, image.PixelHeight));
                                //page_.Width = image.PixelWidth;
                                //page_.Height = image.PixelHeight;

                                XGraphics gfx = XGraphics.FromPdfPage(page_, XGraphicsPdfPageOptions.Append);
                                gfx.SmoothingMode = XSmoothingMode.HighQuality;
                                gfx.DrawImage(image, 0, 0, image.PixelWidth, image.PixelHeight);
                                gfx.Dispose();
                            }
                            ///////////////////////
                        }

                        var percent_ = (float)(pageIdx + 1) * 50 / pageCount + 50;
                        bgw.ReportProgress((int)percent_);
                    }

                    output.Flush();

                    document.Save(output, true);
                }
            }

        }

        private void DoConvertPdfToPdf(BackgroundWorker bgw, ConvertPdfToImageBackgroundModel model_)
        {

            using (var document = PdfDocument.Load(model_.pdfFilePath))
            {
                Directory.CreateDirectory(model_.imageOutputFolderPath);

                int pageCount = document.PageCount;
                MemoryStream[] streamPages = new MemoryStream[pageCount];

                for (int pageIdx = 0; pageIdx < pageCount; pageIdx++)
                {
                    using (var image = document.Render(pageIdx, 300, 300, PdfRenderFlags.CorrectFromDpi))
                    {
                        MemoryStream currentStreamPage = new MemoryStream();
                        image.Save(stream: currentStreamPage, format: ImageFormat.Png);

                        streamPages[pageIdx] = currentStreamPage;
                    }
                    var percent_ = (float)(pageIdx + 1) * 100 / (pageCount * 2);
                    bgw.ReportProgress((int)percent_);
                }

                SavePagesToOnePdf(bgw, model_, streamPages);
            }
        }
    }
}
