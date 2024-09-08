//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

// TODO: Investigate again VSTO / .NET 2.0 security loading problem, 
//       and look at ExecutionContext.SuppressFlow 
//       More info: http://social.msdn.microsoft.com/forums/en-US/clr/thread/0a48607c-5a27-4d12-8e0f-160daed38ef2
//              and http://msdn.microsoft.com/en-us/magazine/cc163644.aspx (Abortable thread pool.)

using System;
using System.Diagnostics;

namespace ExcelDna.Loader
{
    internal static class ProcessHelper
    {
        private static bool _isInitialized = false;
        private static bool _isRunningOnCluster;
        private static int _processMajorVersion;

        public static bool IsRunningOnCluster
        {
            get
            {
                EnsureInitialized();
                return _isRunningOnCluster;
            }
        }

        public static int ProcessMajorVersion
        {
            get
            {
                EnsureInitialized();
                return _processMajorVersion;
            }
        }

        public static bool SupportsClusterSafe
        {
            get
            {
                return IsRunningOnCluster || (ProcessMajorVersion >= 14);
            }
        }

        private static void EnsureInitialized()
        {
            if (!_isInitialized)
            {
                Process hostProcess = Process.GetCurrentProcess();
                _isRunningOnCluster = !(hostProcess.ProcessName.Equals("EXCEL", StringComparison.InvariantCultureIgnoreCase));
                _processMajorVersion = hostProcess.MainModule.FileVersionInfo.FileMajorPart;

                _isInitialized = true;
            }
        }
    }
}
