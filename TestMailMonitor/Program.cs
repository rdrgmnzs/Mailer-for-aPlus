using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NewMailMonitor;

namespace TestMailMonitor
{
    class Program
    {
        static void Main()
        {
            try
            {
                NewMail MailMonitor = new NewMail();
                MailMonitor.Start();
                Console.WriteLine("Hit return to exit");
                Console.ReadLine();
                MailMonitor.Stop();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
