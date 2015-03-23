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
using cbr2pdf;
using System.Threading;
using System.Collections;
using System.Diagnostics;

namespace CbrToPdf
{
    public partial class GUIInterface : Form, ProgessListener
    {
        public string input_bestand;
        private bool finished = false;
        private bool multipleFolder = false;

        public GUIInterface(string[] args)
        {
            InitializeComponent();
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnApplicationExit);

            if (args.Length == 0)
            {
                DialogResult mf = MessageBox.Show("Press Yes, if you want to convert a whole folder.\n\nPress No, If you only want to convert a single file.", "CBR To PDF Conveter", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                if (mf == DialogResult.Yes)
                {
                    using (FolderBrowserDialog dlg = new FolderBrowserDialog())
                    {
                        dlg.Description = "Select a folder";
                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            multipleFolder = true;
                            input_bestand = dlg.SelectedPath;
                            MessageBox.Show("You selected: " + dlg.SelectedPath);
                        }
                    }
                }
                else
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

                if (multipleFolder)
                {
                    string testFile = input_bestand + "\\123456789folder.lock";
                    testWriteAccess(testFile);
                    label1.Text = "Processing folder '" + input_bestand + "'...";
                }
                else
                {
                    string dir = Path.GetDirectoryName(input_bestand);
                    string testFile = dir + "\\" + Path.GetFileNameWithoutExtension(input_bestand) + ".lock";
                    testWriteAccess(testFile);
                    label1.Text = "Processing '" + Path.GetFileName(input_bestand) + "'...";
                }

                backgroundWorker1.WorkerReportsProgress = true;
                backgroundWorker1.WorkerSupportsCancellation = true;
                backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
                backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
                backgroundWorker1.RunWorkerAsync();
                
            }
        }

        private void testWriteAccess(string testFile)
        {
            try
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(testFile);
                file.Close();

                File.Delete(testFile);
            }
            catch (System.UnauthorizedAccessException)
            {
                MessageBox.Show("No write access to the directory of: " + testFile, "CBR To PDF Conveter", MessageBoxButtons.OK, MessageBoxIcon.Error);

                System.Environment.Exit(1);
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
            catch (System.InvalidOperationException)
            {
                System.Environment.Exit(1);
            }
            System.Environment.Exit(0);
        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (backgroundWorker1 != null)
            {
                
                while (!backgroundWorker1.CancellationPending && finished == false)
                {
                    if (multipleFolder)
                    {
                        string[] filePaths = Directory.GetFiles(input_bestand);
                        ArrayList allthread = new ArrayList();//Thread th[];
                        foreach(string fileName in filePaths){

                            ProcessFile pF = new ProcessFile(fileName);
                            pF.setProgressListener(this);
                            Thread th = new Thread(new ThreadStart(pF.startConvertingFile));
                            allthread.Add(th);
                            th.Start();
                        }
                       // bool allDone = false;
                        foreach (Thread x in allthread)
                        {
                            try
                            {
                                Debug.WriteLine("waiting for thread to be stopped");
                                x.Join();
                            }
                            catch (Exception _) { }
                        }
                        finished = true;

                    }
                    else
                    {
                        ProcessFile pF = new ProcessFile(input_bestand);
                        pF.setProgressListener(this);
                        Thread th = new Thread(new ThreadStart(pF.startConvertingFile));
                        th.Start();
                        try
                        {
                            th.Join();
                        }
                        catch (Exception _) { }
                        finished = true;
                    }
                }
            }
        }

        public void progressUpdate(ProcessFile pwFS, int percentageCompleted)
        {
            backgroundWorker1.ReportProgress(percentageCompleted);
            Console.WriteLine(pwFS.getInputFile() + " has a percentage of: " + percentageCompleted);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;

            if (e.ProgressPercentage == 100)
            {
                label1.Text = "Conversion Completed!";
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 100;
            string inputfile = input_bestand;
            string file = Path.GetFileName(inputfile);
            notifyIcon1.Visible = true;
            if (multipleFolder)
            {
                notifyIcon1.ShowBalloonTip(5000, "Completion", "All the files in folder '" + file + "' are converted to PDF.", ToolTipIcon.Info);
            }
            else
            {
                notifyIcon1.ShowBalloonTip(5000, "Completion", "File '" + file + "' has been converted to PDF.", ToolTipIcon.Info);
            }
            

            System.Environment.Exit(0);
        }
    }
}
