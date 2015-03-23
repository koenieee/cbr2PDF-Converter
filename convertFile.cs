using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace cbr2pdf
{
    class ProcessFile
    {
        string inputFile;
        string outputFile;
        string tempFolder;
        string unlockCodeChilkat = "ZIPT34MB34N_2E76BEB1p39E";
        public static int percentage = 0;

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
            percentage = 10;
        }

        public void clearTempFolder()
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

        public Boolean unzipCbz()
        {
            percentage = 20;
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
                percentage = 30;
                return true;
            }
        }

        public Boolean unrarCbr()
        {
            percentage = 20;
            bool raropen, rarunrar;
            Chilkat.Rar rar = new Chilkat.Rar();
            raropen = rar.Open(inputFile);
            rarunrar = rar.Unrar(tempFolder);
            percentage = 30;
            return raropen && rarunrar;
        }

        public string[] getFileNames()
        {
            percentage = 40;
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
            percentage = 50;
            return images;
        }


        public Boolean createPdf(string[] imageArray)
        {
            percentage = 60;
            PdfDocument doc = new PdfDocument();

            foreach (string oneImage in imageArray)
            {
                PdfPage page = doc.AddPage();

                XGraphics xgr = XGraphics.FromPdfPage(page);
                XImage img = XImage.FromFile(oneImage);

                xgr.DrawImage(img, 0, 0, 595, 841); //a normal A4 paper
                
            }
            percentage = 70;
            try
            {
                doc.Save(outputFile);
                doc.Close();
                percentage = 80;
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
                Console.WriteLine("Unrarring");
                if (createPdf(getFileNames()))
                {
                    percentage = 100;
                    Console.WriteLine("Completed");
                }
                else
                {
                    Console.WriteLine("Something went wrong");
                    return;
                }
            }
            else if (unzipCbz() == true)
            {
                Console.WriteLine("Unzipping");
                if (createPdf(getFileNames()))
                {
                    percentage = 100;
                    Console.WriteLine("Completed");
                }
                else
                {
                    Console.WriteLine("Something went wrong.");
                    return;
                }
            }
            else
            {
                Console.WriteLine("Something went complety wrong");
                return;
            }
        }

        public int percentageCompleted()
        {
            return percentage;
        }

    }
}
