using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Settings
{
    public class Config
    {
        private static Configuration init()
        {
            string APPDATA = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

            ConfigurationFileMap fileMap = new ConfigurationFileMap(APPDATA + @"\APMI\APMI.xml"); //Path to your config file
            Configuration configuration = ConfigurationManager.OpenMappedMachineConfiguration(fileMap);
            return configuration;
        }

        /// <summary>
        /// Set aPlus Mail Path.
        /// </summary>
        /// <param name="MailPath">aPlus Mail Path.</param>
        public static void setAPlusMailPath(string MailPath)
        {
            try
            {
                Configuration configuration = init();
                configuration.AppSettings.Settings["APlusMailPath"].Value = MailPath;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get aPlus Mail Path.
        /// </summary>
        /// <returns>Returns aPlus Mail Path.</returns>
        public static string getAPlusMailPath()
        {
            try
            {
                Configuration configuration = init();
                return configuration.AppSettings.Settings["APlusMailPath"].Value;
            }
            catch (Exception)
            {
                throw;
            }
        }



        /// <summary>
        /// Set the IP address for the AS400 to connect to.
        /// </summary>
        /// <param name="IP">AS400 IP address.</param>
        public static void setAS400IP(string IP)
        {
            try
            {
                Configuration configuration = init();
                configuration.AppSettings.Settings["AS400IPaddress"].Value = IP;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get the IP address for the AS400 to connect to.
        /// </summary>
        /// <returns>Returns AS400 IP address.</returns>
        public static string getAS400IP()
        {
            try
            {
                Configuration configuration = init();
                return configuration.AppSettings.Settings["AS400IPaddress"].Value;
            }
            catch (Exception)
            {
                throw;
            }
        }



        /// <summary>
        /// Set aPlus user that has access to UML path.
        /// </summary>
        /// <param name="user">aPlus user name.</param>
        public static void setAPlusUser(string user)
        {
            try
            {
                Configuration configuration = init();
                configuration.AppSettings.Settings["aplusUser"].Value = user;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get aPlus user that has access to UML path.
        /// </summary>
        /// <returns>Returns aPlus user name.</returns>
        public static string getAPlusUser()
        {
            try
            {
                Configuration configuration = init();
                return configuration.AppSettings.Settings["aplusUser"].Value;
            }
            catch (Exception)
            {
                throw;
            }
        }



        /// <summary>
        /// Set aPlus user password.
        /// </summary>
        /// <param name="password">aPlus user password.</param>
        public static void setAPlusPassword(string password)
        {
            try
            {
                Configuration configuration = init();
                string encryptedPass = PasswordEncoder.EncryptWithByteArray(password);
                configuration.AppSettings.Settings["aplusPassword"].Value = encryptedPass;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get aPlus user password.
        /// </summary>
        /// <returns>Returns aPlus user password.</returns>
        public static string getAPlusPassword()
        {
            try
            {
                Configuration configuration = init();
                return PasswordEncoder.DecryptWithByteArray(
                    configuration.AppSettings.Settings["aplusPassword"].Value);
                
            }
            catch (Exception)
            {
                throw;
            }
        }



        /// <summary>
        /// Set interval to way to check for new mail.
        /// </summary>
        /// <param name="timerInterval">Time to way between checking for new mail</param>
        public static void setTimerInterval(string timerInterval)
        {
            try
            {
                Configuration configuration = init();
                configuration.AppSettings.Settings["TimerInterval"].Value = timerInterval;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get interval to way to check for new mail.
        /// </summary>
        /// <returns>Returns TimerInterval.</returns>
        public static string getTimerInterval()
        {
            try
            {
                Configuration configuration = init();
                return configuration.AppSettings.Settings["TimerInterval"].Value;
            }
            catch (Exception)
            {
                throw;
            }
        }



        /// <summary>
        /// Set SMTP Server Adress.
        /// </summary>
        /// <param name="SMTPServer">SMTP Server IP or FQDN</param>
        public static void setSMTPServer(string SMTPServer)
        {
            try
            {
                Configuration configuration = init();
                configuration.AppSettings.Settings["SMTPAddress"].Value = SMTPServer;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get SMTP Server Adress.
        /// </summary>
        /// <returns>Returns SMTP Server adress.</returns>
        public static string getSMTPServer()
        {
            try
            {
                Configuration configuration = init();
                return configuration.AppSettings.Settings["SMTPAddress"].Value;
            }
            catch (Exception)
            {
                throw;
            }
        }



        /// <summary>
        /// Set SMTP Server Port.
        /// </summary>
        /// <param name="SMTPServer">SMTP Server port.</param>
        public static void setSMTPPort(string port)
        {
            try
            {
                Configuration configuration = init();
                configuration.AppSettings.Settings["SMTPPort"].Value = port;
                configuration.Save();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get SMTP Server Port.
        /// </summary>
        /// <returns>Returns SMTP Server port.</returns>
        public static string getSMTPPort()
        {
            try
            {
                Configuration configuration = init();
                return configuration.AppSettings.Settings["SMTPPort"].Value;
            }
            catch (Exception)
            {
                throw;
            }
        }

    }
}
