using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net.Mime;

using Settings;

namespace Mail_Handler
{
    public class mail
    {
        private StreamWriter logFile;

        public List<string> sendTo = new List<string>();
        public List<string> sendCC = new List<string>();
        public List<string> sendBCC = new List<string>();
        public string sendFrom;
        public string sendSubject;
        public string sendMessage;
        public List<string> attachments = new List<string>();

        public mail(StreamWriter logFile)
        {
            this.logFile = logFile;
        }

        /// <summary>
        /// Transmit an email message with
        /// attachments
        /// </summary>
        /// <param name="sendTo">Recipient Email Address</param>
        /// <param name="sendFrom">Sender Email Address</param>
        /// <param name="sendSubject">Subject Line Describing Message</param>
        /// <param name="sendMessage">The Email Message Body</param>
        /// <param name="attachments">A string array pointing to the location of each attachment</param>
        /// <returns>Status Message as String</returns>
        public void SendMessage()
        {
            try
            {
                // Create the basic message
                MailMessage message = new MailMessage();

                // Set e-mail adress message will be send From
                message.From = new MailAddress(sendFrom);

                // Set e-mail adress for all To recepients.
                logFile.WriteLine("Sending mail to:");
                foreach (string mailRecepient in sendTo)
                {
                    logFile.WriteLine(mailRecepient);

                    message.To.Add(mailRecepient);
                }

                // Set e-mail adress for all CC recepients.
                foreach (string mailRecepient in sendCC)
                {
                    logFile.WriteLine(mailRecepient);

                    message.CC.Add(mailRecepient);
                }

                // Set e-mail adress for all CC recepients.
                foreach (string mailRecepient in sendBCC)
                {
                    logFile.WriteLine(mailRecepient);

                    message.Bcc.Add(mailRecepient);
                }

                message.Subject = sendSubject;
                message.Body = sendMessage;
                
                // The attachments arraylist should point to a file location where
                // the attachment resides - add the attachments to the message
                foreach (string attach in attachments)
                {
                    Attachment attached = new Attachment(attach, MediaTypeNames.Application.Octet);
                    message.Attachments.Add(attached);
                }

                // create smtp client at mail server location
                SmtpClient client = new SmtpClient(Config.getSMTPServer());

                // Add credentials
                client.UseDefaultCredentials = true;

                // send message
                client.Send(message);

                foreach (Attachment attach in message.Attachments)
                {
                    attach.Dispose();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                logFile.WriteLine("SendMessage() failed.");
                logFile.WriteLine("Error: " + ex.Message);
                throw;
            }
        }
    }
}
