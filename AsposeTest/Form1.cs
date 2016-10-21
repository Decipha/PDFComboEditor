using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace AsposeTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // Create a PDF license object
            Aspose.Pdf.License license = new Aspose.Pdf.License();
            // Instantiate license file
            license.SetLicense("AsposeTest.Aspose.Total.lic");
            // Set the value to indicate that license will be embedded in the application
            license.Embedded = true;

        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            if (pdfFolderTextBox.Text.Length > 0)
                folderBrowserDialog1.SelectedPath = pdfFolderTextBox.Text;

            DialogResult res = folderBrowserDialog1.ShowDialog(this);
            if (res == DialogResult.OK)
            {
                pdfFolderTextBox.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        FileDetails _fileDetails = null;
        private void processButton_Click(object sender, EventArgs e)
        {
            DirectoryInfo d = new DirectoryInfo(pdfFolderTextBox.Text);
            _fileDetails = new FileDetails( d.GetFiles("*.pdf", SearchOption.TopDirectoryOnly));
            dataGridView1.DataSource = _fileDetails;

            EnableControls(false);

            backgroundWorker1.RunWorkerAsync();
        }

        private void EnableControls(bool enable)
        {
            browseButton.Enabled = enable;
            pdfFolderTextBox.Enabled = enable;
            processButton.Enabled = enable;
            progressBar1.Value = 0;
        }



        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            float counter = 0;
            int total = _fileDetails.Count;
            foreach (FileDetail pdf in _fileDetails)
            {
                try
                {
                    pdf.OutputSize = ProcessFile(pdf.FileName);
                    pdf.Processed = true;
                    backgroundWorker1.ReportProgress((int)(++counter / total * 100));

                }
                catch (Exception ex)
                {


                    string sourceFolder = Path.GetDirectoryName(pdf.FileName);
                    string outputFolder = Path.Combine(sourceFolder, "Converted");

                    if (!Directory.Exists(outputFolder))
                        Directory.CreateDirectory(outputFolder);

                    string outputname = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(pdf.FileName));


                    outputname += ".log";

                    if (File.Exists(outputname))
                        File.AppendAllText(outputname, Environment.NewLine + ex.ToString());
                    else
                        File.WriteAllText(outputname, ex.ToString());
                }
            }

        }


        private long ProcessFile(string filename)
        {
            string inputFile = filename; // Path.Combine(pdfFolderTextBox.Text, filename);
            using (Aspose.Pdf.Document d = new Aspose.Pdf.Document(inputFile))
            {

                d.Info.Title = "Title";
                d.Info.Author = "Author";
                d.Info.Add("scanDocumentID", "ScanDocumentID");
                d.Info.Add("scanBatchID", "ScanBatchID");
                d.Info.Add("scanEnvelopeID", "ScanEnvelopeID");
                d.Info.Add("documentReference", "DocumentRef");
                d.Info.Add("communicationType", "CommunicationType");
                d.Info.Add("documentName", "DocumentName");
                d.Info.Add("agent", "Agent");

                string keywords = string.Empty;
                foreach (string key in d.Info.Keys)
                {
                    if (!"Title,Author,CreationDate,ModDate".Contains(key))
                        keywords += string.Format("{0}={1},", key, d.Info[key]);
                }
                d.Info.Keywords = keywords.TrimEnd(',');

                var before = GC.GetTotalMemory(false);
                GC.Collect();
                var after = GC.GetTotalMemory(false);
                Console.WriteLine("Mem Before Collect: {0}, After {1}", before, after);


                d.Convert("log.xml", Aspose.Pdf.PdfFormat.PDF_A_1B, Aspose.Pdf.ConvertErrorAction.Delete);

                string sourceFolder = Path.GetDirectoryName(inputFile);
                string outputFolder = Path.Combine(sourceFolder, "Converted");

                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                string outputname = Path.Combine(outputFolder, Path.GetFileName(filename));
                d.Save(outputname);

                return new FileInfo(outputname).Length;
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            dataGridView1.Refresh();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            EnableControls(true);
            MessageBox.Show("Processing Complete");
        }




    }
}
