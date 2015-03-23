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

        public ProcessFile(string ipF)
        {
            inputFile = ipF;
            outputFile = Path.GetDirectoryName(inputFile) + "\\" + Path.GetFileNameWithoutExtension(inputFile) + ".pdf";
            tempFolder = System.IO.Path.GetTempPath() + "\\strips";

            if (!Directory.Exists(tempFolder))
            {
                DirectoryInfo di = Directory.CreateDirectory(tempFolder);
            }
            else
            {
                clearTempFolder();
            }
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
                return true;
            }
        }

        public Boolean unrarCbr()
        {
            bool raropen, rarunrar;
            Chilkat.Rar rar = new Chilkat.Rar();
            raropen = rar.Open(inputFile);
            rarunrar = rar.Unrar(tempFolder);
            return raropen && rarunrar;
        }

        public Boolean createPdf()
        {
            PdfDocument doc = new PdfDocument();

            foreach (string source in plaatjes)
            {
                PdfPage page = doc.AddPage();

                XGraphics xgr = XGraphics.FromPdfPage(page);
                XImage img = XImage.FromFile(source);


                xgr.DrawImage(img, 0, 0, 595, 841);

            }
            doc.Save(outputFile);
            doc.Close();
        }

    }
}
