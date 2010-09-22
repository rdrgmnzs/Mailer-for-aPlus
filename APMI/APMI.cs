using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using NewMailMonitor;

namespace APMI
{
    public partial class APMI : ServiceBase
    {
        NewMail MailMonitor;

        public APMI()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            MailMonitor = new NewMail();
            //eventLog.WriteEntry("OnStart");
            MailMonitor.Start();
        }

        protected override void OnStop()
        {
            MailMonitor.Stop();
            //eventLog.WriteEntry("OnStop");
        }
    }
}
