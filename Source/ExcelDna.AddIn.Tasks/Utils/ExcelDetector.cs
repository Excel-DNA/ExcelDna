using System;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using Microsoft.Win32;
using ExcelDna.AddIn.Tasks.Logging;

namespace ExcelDna.AddIn.Tasks.Utils
{
    internal class ExcelDetector : IExcelDetector
    {
        private readonly IBuildLogger _log;

        public ExcelDetector(IBuildLogger log)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public bool TryFindLatestExcel(out string excelExePath)
        {
            excelExePath = null;

            _log.Debug("Trying to find latest version of Excel");

            var versions = (ExcelVersions[])Enum.GetValues(typeof(ExcelVersions));
            var versionsNumbersDescending = versions.Select(v => (int)v).OrderByDescending(vn => vn);

            foreach (var versionNumber in versionsNumbersDescending)
            {
                _log.Debug($"Trying to find {versionNumber} installed");

                var keyPath = $@"Software\Microsoft\Office\{versionNumber}.0\Excel\InstallRoot";

                if (!TryGetExcelExePathFromRegistry(keyPath, out excelExePath)) continue;

                return true;
            }

            return false;
        }

        public bool TryFindExcelBitness(string excelExePath, out Bitness bitness)
        {
            bitness = Bitness.Unknown;

            if (!File.Exists(excelExePath))
            {
                throw new Exception("Excel path specified in Registry not found on disk: " + excelExePath);
            }

            using (var fileStream = File.OpenRead(excelExePath))
            {
                using (var reader = new BinaryReader(fileStream))
                {
                    // See http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx
                    // Offset to PE header is always at 0x3C.
                    // The PE header starts with "PE\0\0" =  0x50 0x45 0x00 0x00,
                    // followed by a 2-byte machine type field (see the document above for the enum).

                    fileStream.Seek(0x3c, SeekOrigin.Begin);
                    var peOffset = reader.ReadInt32();

                    fileStream.Seek(peOffset, SeekOrigin.Begin);
                    var peHead = reader.ReadUInt32();

                    if (peHead != 0x00004550) // "PE\0\0", little-endian
                    {
                        throw new Exception("Unable to find PE header in file");
                    }

                    var machineType = (MachineType)reader.ReadUInt16();

                    switch (machineType)
                    {
                        case MachineType.ImageFileMachineI386:
                            {
                                bitness = Bitness.Bit32;
                                return true;
                            }
                        case MachineType.ImageFileMachineAmd64:
                        case MachineType.ImageFileMachineIa64:
                            {
                                bitness = Bitness.Bit64;
                                return true;
                            }
                        default:
                            {
                                bitness = Bitness.Unknown;
                                return false;
                            }
                    }
                }
            }
        }

        private bool TryGetExcelExePathFromRegistry(string keyPath, out string excelExePath)
        {
            if (TryGetExcelExePathFromRegistry(keyPath, RegistryView.Registry64, out excelExePath))
            {
                _log.Debug($"Found Excel path on {RegistryView.Registry64}: {excelExePath}");
                return true;
            }

            if (TryGetExcelExePathFromRegistry(keyPath, RegistryView.Registry32, out excelExePath))
            {
                _log.Debug($"Found Excel path on {RegistryView.Registry32}: {excelExePath}");
                return true;
            }

            _log.Debug("Unable to find Excel installation path");
            return false;
        }

        private bool TryGetExcelExePathFromRegistry(string keyPath, RegistryView registryView, out string excelExePath)
        {
            excelExePath = null;

            using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
            {
                using (var excelKey = baseKey.OpenSubKey(keyPath, RegistryKeyPermissionCheck.ReadSubTree,
                    RegistryRights.ReadKey))
                {
                    if (excelKey == null) return false;

                    var installRoot = excelKey.GetValue("Path");
                    if (installRoot != null)
                    {
                        excelExePath = Path.Combine(installRoot.ToString(), "EXCEL.EXE");
                        return true;
                    }
                }
            }

            return false;
        }

        private enum ExcelVersions
        {
            // ReSharper disable UnusedMember.Local
            Excel2003 = 11, // Office 2003 - 11.0
            Excel2007 = 12, // Office 2007 - 12.0
            Excel2010 = 14, // Office 2010 - 14.0 (sic!)
            Excel2013 = 15, // Office 2013 - 15.0
            Excel2016 = 16, // Office 2016 - 16.0
            // ReSharper restore UnusedMember.Local
        }

        private enum MachineType : ushort
        {
            ImageFileMachineAmd64 = 0x8664,
            ImageFileMachineI386 = 0x14c,
            ImageFileMachineIa64 = 0x200,
        }
    }
}
