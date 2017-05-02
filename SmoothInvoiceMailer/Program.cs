using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;


using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Google.Apis.Drive.v3.FilesResource;

namespace SmoothInvoiceMailer
{
    class Program
    {

        static void Main(string[] args)
        {
            //run this daily with task scheduler
            //make sure you have your secret files added to the solution
            //create new drive service 
            GoogleDocsHelper driveHelper = new GoogleDocsHelper();
            //create sheets helper
            GoogleSheetHelper gSHelper = new GoogleSheetHelper("smoothInvoiceMailer", "sheets_id.json", "smoothInvoiceMailer");

          
            string templateID = "your template id";

            //get the last invoice date
            DateTime lastinvoiceDate;
            using (StreamReader sr = System.IO.File.OpenText("LastInvoiceDate.txt"))
            {
                string s = sr.ReadToEnd();
                lastinvoiceDate = DateTime.Parse(s);
                //you then have to process the string
            }

            //if last invoice date is exactly 2 weeks ago create a new invoice
            if (lastinvoiceDate.Date.ToShortDateString() == DateTime.Now.AddDays(-14).ToShortDateString())
            {
                Console.WriteLine("Its been 2 weeks create invoice");
            }
            else
            {
                //dont create an invoice and exit
                return;
            }

          
            //Copy template to new file
            Google.Apis.Drive.v3.Data.File newFile = driveHelper.CopyFile( templateID, "GreenSaver Invoice " + InvoiceHelper.CreateInvoiceNumber());


            /*update the cells on the new sheet */
            //update date cell at E3
            gSHelper.UpdateCellValue("E", 3, 5, DateTime.Now.ToShortDateString(), newFile.Id);
            //update inovice number at e5
            gSHelper.UpdateCellValue("E", 5, 5, InvoiceHelper.CreateInvoiceNumber(), newFile.Id);
            //update 1st line item
            string lineItemText = String.Format("new line Item",DateTime.Now.AddDays(-14).ToShortDateString(), DateTime.Now.AddDays(-10).ToShortDateString());
            gSHelper.UpdateCellValue("B", 18, 5, lineItemText, newFile.Id);
            //update 2nd line item
            string lineItemText2 = String.Format("new line litem", DateTime.Now.AddDays(-7).ToShortDateString(), DateTime.Now.AddDays(-3).ToShortDateString());
            gSHelper.UpdateCellValue("B", 20, 5, lineItemText2, newFile.Id);

            //download it as a PDF to temp folder
            string pdfFile = driveHelper.DownloadPDF(newFile.Id, "New Invoice " + InvoiceHelper.CreateInvoiceNumber() + ".pdf");

            //email it to sandy

            GmailHelper gmailHelper = new GmailHelper();
            gmailHelper.SendEmail("to email", "from email", "cc email", "Invoice" + InvoiceHelper.CreateInvoiceNumber(), String.Format(@"body",InvoiceHelper.CreateInvoiceNumber()),"password for gmail", pdfFile);

            //update the last invoice date file
            FileInfo fi = new FileInfo(@"LastInvoiceDate.txt");
            using (StreamWriter sw = new StreamWriter(fi.Open(FileMode.Truncate))) 
            {
                sw.Write(DateTime.Now.ToShortDateString());
            }

            //updateFileToDrive
            driveHelper.UploadFileToDrive(pdfFile, "folder ID");
            //print doucment
            SendToPrinter(pdfFile);
            //delete temp file
        //    System.IO.File.Delete(pdfFile);
            

        }


        private static void SendToPrinter(string file)
        {
            ProcessStartInfo info = new ProcessStartInfo();
            info.Verb = "print";
            info.FileName = file;
            info.CreateNoWindow = true;
           // info.WindowStyle = ProcessWindowStyle.Hidden;

            Process p = new Process();
            p.StartInfo = info;
            p.Start();

            p.WaitForInputIdle();
            System.Threading.Thread.Sleep(3000);
            if (false == p.CloseMainWindow())
                p.Kill();
        }

    }

   


}
