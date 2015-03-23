using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace cbr2pdf
{
    public class ProcessFile
    {
        string inputFile;
        string outputFile;
        string tempFolder;
        string unlockCodeChilkat = "ZIPT34MB34N_2E76BEB1p39E";
        ProgessListener whoIsListener;
        

        public ProcessFile(string ipF) //process CBR or CBZ file. call Constructor and then startConvertingFiles()
        {
            inputFile = ipF;
            outputFile = Path.GetDirectoryName(inputFile) + "\\" + Path.GetFileNameWithoutExtension(inputFile) + ".pdf";
            tempFolder = System.IO.Path.GetTempPath() + "\\strips_"+Path.GetFileName(inputFile).Replace(" ","_");

            if (!Directory.Exists(tempFolder))
            {
                DirectoryInfo di = Directory.CreateDirectory(tempFolder);
            }
            else
            {
                clearTempFolder();
            }
            //Updatepercentage = 10;
        }

        public void setProgressListener(ProgessListener ls)
        {
            whoIsListener = ls;
            whoIsListener.progressUpdate(this, 10); 
        }

        public string getInputFile()
        {
            return Path.GetFileNameWithoutExtension(inputFile);
        }

        private void clearTempFolder()
        {
            System.IO.DirectoryInfo temp = new DirectoryInfo(tempFolder);

            foreach (FileInfo file in temp.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in temp.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        private Boolean unzipCbz()
        {
            whoIsListener.progressUpdate(this, 20);
            bool zipopen;
            int zipunzip;
            Chilkat.Zip zip = new Chilkat.Zip();
            zip.UnlockComponent(unlockCodeChilkat);

            zipopen = zip.OpenZip(inputFile);
            zipunzip = zip.Unzip(tempFolder);
            if (zipunzip == -1 || zipopen == false)
            {
                return false;
            }
            else
            {
                whoIsListener.progressUpdate(this, 30);
                return true;
            }
        }

        private Boolean unrarCbr()
        {
            whoIsListener.progressUpdate(this, 20);
            bool raropen, rarunrar;
            Chilkat.Rar rar = new Chilkat.Rar();
            raropen = rar.Open(inputFile);
            rarunrar = rar.Unrar(tempFolder);
            whoIsListener.progressUpdate(this, 30);
            return raropen && rarunrar;
        }

        private string[] getFileNames()
        {
            whoIsListener.progressUpdate(this, 40);
            string[] folders_all = System.IO.Directory.GetDirectories(tempFolder, "*", System.IO.SearchOption.AllDirectories);

            string[] images;
            if (folders_all.Length == 0)
            {
                images = Directory.GetFiles(tempFolder, "*.jpg");
            }
            else
            {
                string folder = folders_all[0];
                images = Directory.GetFiles(folder, "*.jpg");
            }
            whoIsListener.progressUpdate(this, 50);
            return images;
        }


        private Boolean createPdf(string[] imageArray)
        {
            whoIsListener.progressUpdate(this, 70);
            PdfDocument doc = new PdfDocument();

            foreach (string oneImage in imageArray)
            {
                PdfPage page = doc.AddPage();

                XGraphics xgr = XGraphics.FromPdfPage(page);
                XImage img = XImage.FromFile(oneImage);

                xgr.DrawImage(img, 0, 0, 595, 841); //a normal A4 paper
                img.Dispose();
                
            }
            whoIsListener.progressUpdate(this, 80);
            try
            {
                doc.Save(outputFile);
                doc.Close();
                whoIsListener.progressUpdate(this, 90);
                return true;
            }
            catch(Exception _){
                return false;
            }

        }

        public void startConvertingFile()
        {
            if (unrarCbr() == true)
            {
                Debug.WriteLine("Unrarring");
                if (createPdf(getFileNames()))
                {
                    clearTempFolder();
                    whoIsListener.progressUpdate(this, 100);
                    Debug.WriteLine("Completed");
                }
                else
                {
                    Debug.WriteLine("Something went wrong");
                    return;
                }
            }
            else if (unzipCbz() == true)
            {
                Debug.WriteLine("Unzipping");
                if (createPdf(getFileNames()))
                {
                    clearTempFolder();
                    whoIsListener.progressUpdate(this, 100);
                    Debug.WriteLine("Completed");
                }
                else
                {
                    Debug.WriteLine("Something went wrong.");
                    return;
                }
            }
            else
            {
                Debug.WriteLine("Something went complety wrong");
                return;
            }
        }
    }
}
