using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Timers;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Reflection;

using Mail_Handler;
using Settings;

namespace NewMailMonitor
{
    public class NewMail
    {
        private Thread mailMonitorThread;

        private static System.Timers.Timer aTimer;
        private static int timerInterval = int.Parse(Config.getTimerInterval());

        private static string AS400IP = Config.getAS400IP();
        private static string userName = Config.getAPlusUser();
        private static string userPassword = Config.getAPlusPassword();

        ImpersonateUser user = new ImpersonateUser();

        private static string mailFolder = Config.getAPlusMailPath() + @"MSHICTL\";
        private static string attachFolder = Config.getAPlusMailPath() + @"MSATTCH\";

        private static string APPDATA = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        private static StreamWriter logFile = new StreamWriter(APPDATA + @"\APMI\log.txt");

        public void Start()
        {
            //Make log write in real time
            logFile.AutoFlush = true;

            //Impersonates User to log on to AS400
            user.Impersonate(userName, AS400IP, userPassword);

            mailMonitorThread = new Thread(GetMailTimerTread);
            mailMonitorThread.IsBackground = true;
            mailMonitorThread.Name = "MailMonitor";
            mailMonitorThread.Start();
        }
        public void Stop()
        {
            mailMonitorThread.Abort();
        }

        /// <summary> 
        /// Set and starts the mail checker, with a 
        /// timed interval to check for e-mail
        /// </summary>
        private static void GetMailTimerTread()
        {
            try
            {
                // Create a timer interval with a number of second
                // stored in the settings file.
                aTimer = new System.Timers.Timer(timerInterval * 1000);

                // Hook up the Elapsed event for the timer.
                aTimer.Elapsed += new ElapsedEventHandler(GetMailFromDirectory);
                aTimer.Enabled = true;

                // If the timer is declared in a long-running method, use
                // KeepAlive to prevent garbage collection from occurring
                // before the method ends.
                GC.KeepAlive(aTimer);

                logFile.WriteLine("Check Mail timer started.");
            }
            catch (Exception ex)
            {
                logFile.WriteLine("Check Mail timer filed start.");
                logFile.WriteLine("Error: " + ex.Message);
                throw;
            }
        }

        /// <summary> 
        /// Get the e-mail files from the directory 
        /// and processes it into an e-mail.
        /// </summary>
        private static void GetMailFromDirectory(object source, ElapsedEventArgs e)
        {
            try
            {
                mail newMail = new mail(logFile);

                DirectoryInfo Dir = new DirectoryInfo(mailFolder);
                FileInfo[] FileList = Dir.GetFiles("*.*", SearchOption.AllDirectories);

                //Process all new e-mail files.
                foreach (FileInfo FI in FileList)
                {
                    string s = "";

                    Console.WriteLine(FI.FullName);

                    using (StreamReader sr = FI.OpenText())
                    {

                        //Process the text file into and email.
                        while ((s = sr.ReadLine()) != null)
                        {
                            string[] tempString = s.Split(':');

                            string mailNumber = FI.Name.Substring(1);

                            switch (tempString[0])
                            {
                                case "ATCH":
                                    string attchmentName = (tempString[1].TrimStart(' ')).TrimEnd(' ') + ".txt";

                                    File.Move(attachFolder + "A" + mailNumber, attachFolder + attchmentName);
                                    newMail.attachments.Add(attachFolder + attchmentName);
                                    break;

                                case "FROM":
                                    newMail.sendFrom = (tempString[1].TrimStart(' ')).TrimEnd(' ');
                                    break;

                                case "TO":
                                    newMail.sendTo.Add((tempString[1].TrimStart(' ')).TrimEnd(' '));
                                    break;

                                case "CC":
                                    newMail.sendCC.Add((tempString[1].TrimStart(' ')).TrimEnd(' '));
                                    break;

                                case "BCC":
                                    newMail.sendBCC.Add((tempString[1].TrimStart(' ')).TrimEnd(' '));
                                    break;

                                case "SUBJ":
                                    newMail.sendSubject = (tempString[1].TrimStart(' ')).TrimEnd(' ');
                                    break;

                                case "TEXT":
                                    newMail.sendMessage = (tempString[1].TrimStart(' ')).TrimEnd(' ');
                                    break;

                                default:
                                    Console.WriteLine("NOTHING");
                                    break;
                            }

                        }
                    }

                    logFile.WriteLine("Sending Mail: " + FI.Name.Substring(1));
                    newMail.SendMessage();

                    foreach (string file in newMail.attachments)
                    {
                        FileInfo attch = new FileInfo(file);
                        File.SetAttributes(file, FileAttributes.Normal);
                        attch.Delete();
                    }

                    newMail = new mail(logFile);

                    FI.Delete();
                }
            }
            catch (Exception ex)
            {
                logFile.WriteLine("Get Mail From Directory (AS400 IFS) failed.");
                logFile.WriteLine("Error: " + ex.Message);
                throw;
            }
        }
  
    }
}
