using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Outlook;

namespace Outlook_send_message
{
    class Program
    {
        static Application app = new Application();
        static string subject = "Subject";
        static AutoResetEvent WaitForReply;
        static string bodyEmail = "What ever first line\r\n" +
                "Whatever Second Line \r\n" +
                "And so on ...\r\n";
        static string toEmail = "SourceEmail";

        static void Main(string[] args)
        {
            WaitForReply = new AutoResetEvent(false);

            //Use this to send with an attachment
            /*
             SendEmailWithAttachment(subject, toEmail, bodyEmail, new List<string> { @"C:\Files\attachment.doc" });
            */

            //Sending the emails
            SendEmail(subject, toEmail, bodyEmail);
           


            //Waiting for a response
            app.NewMailEx += App_NewMailEx;

            //Wait until this thread is relased (which will happen when we receive the reply and delete it
            WaitForReply.WaitOne();
          
        }

        private static void App_NewMailEx(string EntryIDCollection)
        {         
            MailItem newMail = (MailItem)app.Session.GetItemFromID(EntryIDCollection, System.Reflection.Missing.Value);


            //This is a cryptic way to get String.contains("") in case-insensitive way
            if (newMail.Subject.IndexOf(subject, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                newMail.Delete();
                //Shut the application down, you can comment this if you want it to continue for deleting more replies
                WaitForReply.Set();
            }
        }

      public static void SendEmail(string subjectEmail,
             string toEmail, string bodyEmail)
        {       
            MailItem oMsg = app.CreateItem(OlItemType.olMailItem);
            oMsg.DeleteAfterSubmit = true;  //Delete the message from sent box
            oMsg.Subject = subjectEmail;
            oMsg.To = toEmail;
            var ins = oMsg.GetInspector;
            oMsg.HTMLBody = bodyEmail.Replace("\r\n", "<br />") + oMsg.HTMLBody; //Must append oMsg.HTMLBody to get the default signature
            oMsg.Send();
        }


        public static void SendEmailWithAttachment(string subjectEmail, string toEmail, string bodyEmail, List<String> attachments)
        {
            MailItem oMsg = app.CreateItem(OlItemType.olMailItem);
            oMsg.DeleteAfterSubmit = true;
            oMsg.Subject = subjectEmail;
            oMsg.To = toEmail;
            var ins = oMsg.GetInspector;
            foreach (var attachment in attachments)
            {
                oMsg.Attachments.Add(attachment, OlAttachmentType.olByValue);
            }
            oMsg.HTMLBody = bodyEmail.Replace("\r\n", "<br />") + oMsg.HTMLBody; //Must append oMsg.HTMLBody to get the default signature
            oMsg.Send();
        }
    }
}
