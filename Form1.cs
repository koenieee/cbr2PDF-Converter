//Program: Cbr2PDF Converter
//Author: Koen
//Version: 8.0
//License: GPL/GNU
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using Chilkat;
using System.Security;
using System.Security.Permissions;
using System.Security.AccessControl;
using System.Security.Principal;

namespace CbrToPdf
{
    public partial class Form1 : Form
    {
        public string input_bestand;
        private bool finished = false;
        

        public Form1(string[] args)
        {
            InitializeComponent();
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnApplicationExit);

            if (args.Length == 0)
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "CBR files (*.cbr)|*.cbr|CBZ files (*.cbz)|*.cbz";

                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile); ; 
                dialog.Title = "Select a CBR/CBZ  File";
                DialogResult result = dialog.ShowDialog();

                if (result.Equals(DialogResult.Cancel) || result.Equals(DialogResult.Abort))
                {
                    MessageBox.Show("You didn't select a file, exiting CBR to PDF Converter.", "CBR To PDF Conveter", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    System.Environment.Exit(1);
                }
                else if (result.Equals(DialogResult.OK))
                {
                    input_bestand = dialog.FileName;
                }
            }
            else
            {
                input_bestand = args[0];
            }

            if (string.IsNullOrEmpty(input_bestand))
            {
                System.Environment.Exit(1);
            }
            else
            {


                string dir = Path.GetDirectoryName(input_bestand);
                string testFile = dir + "\\" + Path.GetFileNameWithoutExtension(input_bestand) + ".lock"; ;
                try
                {
                    System.IO.StreamWriter file = new System.IO.StreamWriter(testFile);
                    file.Close();

                    File.Delete(testFile);
                }
                catch (System.UnauthorizedAccessException)
                {
                    MessageBox.Show("No write access to this directory: " + dir, "CBR To PDF Conveter", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    System.Environment.Exit(1);
                }

                //backgroundWorker1.RunWorkerAsync();
                // //backgroundWorker1.WorkerReportsProgress = true;
                //backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
                //


                backgroundWorker1.WorkerReportsProgress = true;
                backgroundWorker1.WorkerSupportsCancellation = true;
                backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
                backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
                backgroundWorker1.RunWorkerAsync();
                label1.Text = "Processing '" + Path.GetFileName(input_bestand) + "'...";
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            notifyIcon1.Icon = null;
            notifyIcon1.Dispose();
            Application.DoEvents();
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                backgroundWorker1.CancelAsync();
            }
            catch(System.InvalidOperationException){
                System.Environment.Exit(1);
            }
            System.Environment.Exit(1);
        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (backgroundWorker1 != null)
            {
                while (!backgroundWorker1.CancellationPending && finished == false)
                {
                    int i = 0;
                    backgroundWorker1.ReportProgress(i);

                    //============Variables=========================//
                    string inputfile = input_bestand;
                    

                    string inputfolder = System.IO.Path.GetTempPath() + "\\strips";
                    string outputfile = Path.GetDirectoryName(inputfile) + "\\" + Path.GetFileNameWithoutExtension(inputfile) + ".pdf";

                    i = 20;
                    backgroundWorker1.ReportProgress(i);
                    // Determine whether the directory exists.
                    if (!Directory.Exists(inputfolder))
                    {
                        DirectoryInfo di = Directory.CreateDirectory(inputfolder);
                    }


                    System.IO.DirectoryInfo downloadedMessageInfo = new DirectoryInfo(System.IO.Path.GetTempPath() + "\\strips");

                    foreach (FileInfo file in downloadedMessageInfo.GetFiles())
                    {
                        file.Delete();
                    }
                    foreach (DirectoryInfo dir in downloadedMessageInfo.GetDirectories())
                    {
                        dir.Delete(true);
                    }

                    i = 40;
                    backgroundWorker1.ReportProgress(i);

                    //CBR Unzipping=========================================//
                    Chilkat.Rar rar = new Chilkat.Rar();
                    Chilkat.Zip zip = new Chilkat.Zip();
                    zip.UnlockComponent("ZIPT34MB34N_2E76BEB1p39E");
                    bool rarsuccess;
                    bool zipsucces;

                    rarsuccess = rar.Open(inputfile);
                    if (rarsuccess != true)
                    {
                        zipsucces = zip.OpenZip(inputfile);
                        if (zipsucces != true)
                        {
                            MessageBox.Show(zip.LastErrorText);
                            System.Environment.Exit(1);
                        }
                    }

                    rarsuccess = rar.Unrar(inputfolder);
                    if (rarsuccess != true)
                    {
                        int zipreturn = zip.Unzip(inputfolder);
                        if (zipreturn == -1)
                        {
                            MessageBox.Show(zip.LastErrorText);
                            System.Environment.Exit(1);
                        }
                        
                    }
                    else
                    {

                    }
                    i = 60;
                    backgroundWorker1.ReportProgress(i);


                    string[] folders_all = System.IO.Directory.GetDirectories(inputfolder, "*", System.IO.SearchOption.AllDirectories);

                    string[] plaatjes;
                    if (folders_all.Length == 0)
                    {
                        plaatjes = Directory.GetFiles(inputfolder, "*.jpg");
                    }
                    else
                    {
                        string folder = folders_all[0];
                        plaatjes = Directory.GetFiles(folder, "*.jpg");
                    }

                    PdfDocument doc = new PdfDocument();

                    foreach (string source in plaatjes)
                    {
                        PdfPage page = doc.AddPage();

                        XGraphics xgr = XGraphics.FromPdfPage(page);
                        XImage img = XImage.FromFile(source);


                        xgr.DrawImage(img, 0, 0, 595, 841);

                    }
                    i = 80;
                    backgroundWorker1.ReportProgress(i);

                    
                    doc.Save(outputfile);
                    doc.Close();
                    i = 100;
                    finished = true;
                    backgroundWorker1.ReportProgress(i);
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;

            if (e.ProgressPercentage == 100)
            {
                label1.Text = "Conversion Completed!";
            }
        }

        public void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 100;
            string inputfile = input_bestand;
            string file = Path.GetFileName(inputfile);
            notifyIcon1.Visible = true;

            notifyIcon1.ShowBalloonTip(5000, "Completion", "File '" + file + "' has been converted to PDF.", ToolTipIcon.Info);
            
            System.Environment.Exit(1);
        }
    }
}
