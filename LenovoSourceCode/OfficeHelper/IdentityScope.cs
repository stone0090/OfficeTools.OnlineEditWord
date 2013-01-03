using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace LenovoCW.MOA
{
    public class IdentityScope : IDisposable
        {
            // obtains user token   
            [DllImport("advapi32.dll", SetLastError = true)]
            static extern bool LogonUser(string pszUsername, string pszDomain, string pszPassword,
                int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

            // closes open handes returned by LogonUser   
            [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            extern static bool CloseHandle(IntPtr handle);

            [DllImport("Advapi32.DLL")]
            static extern bool ImpersonateLoggedOnUser(IntPtr hToken);

            [DllImport("Advapi32.DLL")]
            static extern bool RevertToSelf();
            const int LOGON32_PROVIDER_DEFAULT = 0;
            const int LOGON32_LOGON_NEWCREDENTIALS = 9;//域控中的需要用:Interactive = 2   
            private bool disposed;
            public  IdentityScope(string sUsername, string sDomain, string sPassword)
            {
                // initialize tokens   
                IntPtr pExistingTokenHandle = new IntPtr(0);
                IntPtr pDuplicateTokenHandle = new IntPtr(0);

                try
                {
                    // get handle to token   
                    bool bImpersonated = LogonUser(sUsername, sDomain, sPassword,
                        LOGON32_LOGON_NEWCREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref pExistingTokenHandle);

                    if (true == bImpersonated)
                    {
                        if (!ImpersonateLoggedOnUser(pExistingTokenHandle))
                        {
                            int nErrorCode = Marshal.GetLastWin32Error();
                            throw new Exception("ImpersonateLoggedOnUser error;Code=" + nErrorCode);
                        }
                    }
                    else
                    {
                        int nErrorCode = Marshal.GetLastWin32Error();
                        throw new Exception("LogonUser error;Code=" + nErrorCode);
                    }
                }
                finally
                {
                    // close handle(s)   
                    if (pExistingTokenHandle != IntPtr.Zero)
                        CloseHandle(pExistingTokenHandle);
                    if (pDuplicateTokenHandle != IntPtr.Zero)
                        CloseHandle(pDuplicateTokenHandle);
                }
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!disposed)
                {
                    RevertToSelf();
                    disposed = true;
                }
            }

            public void Dispose()
            {
                Dispose(true);
            }
        } 

}
