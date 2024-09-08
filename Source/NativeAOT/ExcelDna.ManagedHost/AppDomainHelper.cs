using System;
using System.Diagnostics;
using System.Security;
using System.Security.Permissions;

namespace ExcelDna.ManagedHost
{
    public static class AppDomainHelper
    {
        // This method is called from unmanaged code in a temporary AppDomain, just to be able to call
        // the right AppDomain.CreateDomain overload.
        public static AppDomain CreateFullTrustSandbox()
        {
#if NETFRAMEWORK
            try
            {
                Debug.Print("CreateSandboxAndInitialize - in loader AppDomain with Id: " + AppDomain.CurrentDomain.Id);

                PermissionSet pset = new PermissionSet(PermissionState.Unrestricted);
                AppDomainSetup loaderAppDomainSetup = AppDomain.CurrentDomain.SetupInformation;
                AppDomainSetup sandboxAppDomainSetup = new AppDomainSetup();
                sandboxAppDomainSetup.ApplicationName = loaderAppDomainSetup.ApplicationName;
                sandboxAppDomainSetup.ConfigurationFile = loaderAppDomainSetup.ConfigurationFile;
                sandboxAppDomainSetup.ApplicationBase = loaderAppDomainSetup.ApplicationBase;
                sandboxAppDomainSetup.ShadowCopyFiles = loaderAppDomainSetup.ShadowCopyFiles;
                sandboxAppDomainSetup.ShadowCopyDirectories = loaderAppDomainSetup.ShadowCopyDirectories;

                // create the sandboxed domain
                AppDomain sandbox = AppDomain.CreateDomain(
                    "FullTrustSandbox(" + AppDomain.CurrentDomain.FriendlyName + ")",
                    null,
                    sandboxAppDomainSetup,
                    pset);

                Debug.Print("CreateFullTrustSandbox - sandbox AppDomain created. Id: " + sandbox.Id);

                return sandbox;
            }
            catch (Exception ex)
            {
                Debug.Print("Error during CreateFullTrustSandbox: " + ex.ToString());
                return AppDomain.CurrentDomain;
            }
#else
        return AppDomain.CurrentDomain;
#endif
        }
    }
}
