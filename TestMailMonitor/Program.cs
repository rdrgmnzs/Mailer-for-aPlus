using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NewMailMonitor;

namespace TestMailMonitor
{
    class Program
    {
        // Used to test MailMonitor
        // Quit if any keys are pressed
        // or write any trown exeptions to the command line.
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
