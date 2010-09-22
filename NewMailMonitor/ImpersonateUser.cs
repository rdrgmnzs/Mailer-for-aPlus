using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Reflection;

namespace NewMailMonitor
{
    class ImpersonateUser
    {
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int LogonUser(String lpszUsername, String lpszDomain,
        String lpszPassword,
        int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public extern static bool CloseHandle(IntPtr hToken);

        public void Impersonate(string userName, string domain, string password)
        {

            IntPtr token = IntPtr.Zero;
            IntPtr dupToken = IntPtr.Zero;

            WindowsIdentity tempWindowsIdentity;

            try
            {
                // request default security provider a logon token with LOGON32_LOGON_NEW_CREDENTIALS,
                // token returned is impersonation token, no need to duplicate
                if (LogonUser(userName, domain, password, 9, 0, ref token) != 0)
                {
                    tempWindowsIdentity = new WindowsIdentity(token);
                    impersonationContext = tempWindowsIdentity.Impersonate();
                    // close impersonation token, no longer needed
                    CloseHandle(token);
                }
            }
            catch
            {
                throw;
            }
        }

        WindowsImpersonationContext impersonationContext;
    }
}
