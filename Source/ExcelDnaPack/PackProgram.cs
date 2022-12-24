//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

using System.IO;
using ExcelDna.Integration;
using ExcelDna.AddIn.Tasks.Logging;
using ExcelDna.PackedResources.Logging;

namespace ExcelDnaPack
{
    class PackProgram
    {
        static string usageInfo =
@"ExcelDnaPack Usage
------------------
ExcelDnaPack is a command-line utility to pack ExcelDna add-ins into a single .xll file.

Usage: ExcelDnaPack.exe dnaPath [/O outputPath] [/Y] 

  dnaPath           The path to the primary .dna file for the ExcelDna add-in.
  /Y                If the output .xll exists, overwrite without prompting.
  /NoCompression    no compress (LZMA) of resources
  /NoMultiThreading no multi threading to ensure deterministic packing order
  /O outPath        Output path - default is <dnaPath>-packed.xll.

Example: ExcelDnaPack.exe MyAddins\FirstAddin.dna
		 The packed add-in file will be created as MyAddins\FirstAddin-packed.xll.

The template add-in host file (the copy of ExcelDna.xll renamed to FirstAddin.xll) is 
searched for in the same directory as FirstAddin.dna.

The Excel-Dna integration assembly (ExcelDna.Integration.dll) is searched for 
  1. in the same directory as the .dna file, and if not found there, 
  2. in the same directory as the ExcelDnaPack.exe file.

ExcelDnaPack will also pack the configuration file FirstAddin.xll.config if it is 
found next to FirstAddin.dna.
Other assemblies are packed if marked with Pack=""true"" in the .dna file.
";

        static int Main(string[] args)
        {
            int exitCode;

            try
            {
                exitCode = Pack(args);
            }
            catch (Exception ex)
            {
                exitCode = 1;
                Console.Error.WriteLine(ex.ToString());
            }

            return exitCode;
        }

        private static int Pack(string[] args)
        {
            var buildLogger = new ConsoleLogger(nameof(PackProgram));

            //			string testLib = @"C:\Work\ExcelDna\Version\ExcelDna-0.23\Source\ExcelDnaPack\bin\Debug\exceldna.xll";
            //			ResourceHelper.ResourceLister rl = new ResourceHelper.ResourceLister(testLib);
            //			rl.ListAll();

            //			//ResourceHelper.ResourceUpdater.Test(testLib);
            //			return;

            // Force jit-load of ExcelDna.Integration assembly
            int unused = XlCall.xlAbort;

            if (args.Length < 1)
            {
                Console.Write("No .dna file specified.\r\n\r\n" + usageInfo);
                return 1;
            }

            // Special path when we're building ExcelDna, to pack Loader and Integration
            if (args[0] == "PackXllBuild")
            {
                string xllFullPath = args[1];
                bool includePdb = (args[2] == "Debug");
                PackXllBuild(xllFullPath, includePdb, buildLogger);
                return 0;
            }

            string dnaPath = Path.GetFullPath(args[0]);
            string xllOutputPath = null;
            bool overwrite = false;
            bool compress = true;
            bool multithreading = true;

            // TODO: Replace with an args-parsing routine.
            if (args.Length > 1)
            {
                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].Equals("/O", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (i >= args.Length - 1)
                        {
                            // Too few args.
                            Console.Write("Invalid command-line arguments.\r\n\r\n" + usageInfo);
                            return 1;
                        }
                        xllOutputPath = args[i + 1];
                    }
                    else if (args[i].Equals("/Y", StringComparison.CurrentCultureIgnoreCase))
                    {
                        overwrite = true;
                    }
                    else
                    if (args[i].Equals("/NoCompression", StringComparison.CurrentCultureIgnoreCase))
                    {
                        compress = false;
                    }
                    else
                    if (args[i].Equals("/NoMultiThreading", StringComparison.CurrentCultureIgnoreCase))
                    {
                        multithreading = false;
                    }
                }
            }

            int result = ExcelDna.PackedResources.ExcelDnaPack.Pack(dnaPath, xllOutputPath, compress, multithreading, overwrite, usageInfo, null, false, null, false, null, buildLogger);

#if DEBUG
            if (result == 0)
            {
                Console.WriteLine("Press any key to exit.");
                try
                {
                    Console.ReadKey();
                }
                catch (InvalidOperationException) // Exception when running pack when building in Visual Studio.
                {
                }
            }
#endif

            return result;
        }

        private static void PackXllBuild(string xllFullPath, bool includePdb, IBuildLogger buildLogger)
        {
            ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(xllFullPath, false, buildLogger);

            var xllDir = Path.GetDirectoryName(xllFullPath);
            ru.AddAssembly(Path.Combine(xllDir, "ExcelDna.ManagedHost.dll"), compress: false, multithreading: false, includePdb);
            ru.AddAssembly(Path.Combine(xllDir, "ExcelDna.Loader.dll"), compress: true, multithreading: false, includePdb);
            ru.AddAssembly(Path.Combine(xllDir, "ExcelDna.Integration.dll"), compress: true, multithreading: false, includePdb);
            ru.EndUpdate();
        }
    }
}
