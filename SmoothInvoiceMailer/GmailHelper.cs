using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SmoothInvoiceMailer
{
    public class GmailHelper
    {
        public GmailHelper()
        {


        }

        public void SendEmail(string toEmail, string fromEmail,string ccEmail, string subject,string body,string password,string fileAttachment)
        {
         
            using (MailMessage message = new MailMessage() { From = new MailAddress(fromEmail) })
            {
                //create html message
                message.IsBodyHtml = true;
                //add email addresses
            
                message.To.Add(new MailAddress(toEmail));
                message.CC.Add(new MailAddress(ccEmail));

                //create smpt client and authenticate
                using (SmtpClient client = new SmtpClient())
                {
                    NetworkCredential netCred = new NetworkCredential(fromEmail, password);
                    client.Host = "smtp.gmail.com";
                    client.Port = 587;
                    client.UseDefaultCredentials = false;
                    client.Credentials = netCred;
                    client.EnableSsl = true;
                    //add subject
                    message.Subject = subject;
                    //add body
                    message.Body = body;
                    //create final message
                    //Show the EmailReportChooser Form inside the try

                     Attachment mailAttachment = new Attachment(fileAttachment);
                     message.Attachments.Add(mailAttachment);
                     client.Send(message);
                     Console.WriteLine("Email sent to " + message.To);
                    }

                }
            }
    }
    
}
