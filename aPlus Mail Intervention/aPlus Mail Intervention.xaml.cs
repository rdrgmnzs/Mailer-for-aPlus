using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ServiceProcess;
using Mail_Handler;
using NewMailMonitor;
using Settings;

namespace aPlus_Mail_Intervention
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ServiceController service = new ServiceController("APMI");

        public MainWindow()
        {
            InitializeComponent();
            SMTPServerTxtBox.Text = Config.getSMTPServer();
            SMTPServerPortTxtBox.Text = (Config.getSMTPPort()).ToString();
            aPlusMailTxtBox.Text = Config.getAPlusMailPath();
            MailCheckIntervalTxtBox.Text = (Config.getTimerInterval()).ToString();
            AS400IPTxtBox.Text = Config.getAS400IP();
            userNameTxtBox.Text = Config.getAPlusUser();
            passwordTxtBox.Password = Config.getAPlusPassword();
            
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Stop Service and wait till it has stoped.
                if (service.Status.Equals(ServiceControllerStatus.Running))
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped);
                }

                //Update all app settings.
                Config.setSMTPServer(SMTPServerTxtBox.Text);
                Config.setSMTPPort(SMTPServerPortTxtBox.Text);
                Config.setAPlusMailPath(aPlusMailTxtBox.Text);
                Config.setTimerInterval(MailCheckIntervalTxtBox.Text);

                Config.setAS400IP(AS400IPTxtBox.Text);
                Config.setAPlusUser(userNameTxtBox.Text);
                Config.setAPlusPassword(passwordTxtBox.Password);

                //Start Service and wait till it has started.
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running);

                //Close curent application
                this.Close();
            }
            catch
            {
                throw;
            }
        }

        private void ApplyBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Stop Service and wait till it has stoped.
                if (service.Status.Equals(ServiceControllerStatus.Running))
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped);
                }

                //Update all app settings.
                Config.setSMTPServer(SMTPServerTxtBox.Text);
                Config.setSMTPPort(SMTPServerPortTxtBox.Text);
                Config.setAPlusMailPath(aPlusMailTxtBox.Text);
                Config.setTimerInterval(MailCheckIntervalTxtBox.Text);

                Config.setAS400IP(AS400IPTxtBox.Text);
                Config.setAPlusUser(userNameTxtBox.Text);
                Config.setAPlusPassword(passwordTxtBox.Password);

                //Start Service and wait till it has started.
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running);
            }
            catch
            {
                throw;
            }
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
    }
}
