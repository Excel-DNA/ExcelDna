//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

using System.IO;
using ExcelDna.Integration;

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
                PackXllBuild(xllFullPath, includePdb);
                return 0;
            }

			string dnaPath = Path.GetFullPath(args[0]);
			string dnaDirectory = Path.GetDirectoryName(dnaPath);
//			string dnaFileName = Path.GetFileName(dnaPath);
			string dnaFilePrefix = Path.GetFileNameWithoutExtension(dnaPath);
			string configPath = Path.ChangeExtension(dnaPath, ".xll.config");
            string xllInputPath = Path.Combine(dnaDirectory, dnaFilePrefix + ".xll");
			string xllOutputPath = Path.Combine(dnaDirectory, dnaFilePrefix + "-packed.xll");
            bool overwrite = false;
            bool compress = true;
            bool multithreading = true;

			if (!File.Exists(dnaPath))
			{
				Console.Error.Write("ERROR: Add-in .dna file " + dnaPath + " not found.\r\n\r\n" + usageInfo);
				return 1;
			}

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
                    } else
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

			if (File.Exists(xllOutputPath))
			{
				if (overwrite == false)
				{
					Console.Write("Output .xll file " + xllOutputPath + " already exists. Overwrite? [Y/N] ");
					string response = Console.ReadLine();
					if (!response.Equals("Y", StringComparison.CurrentCultureIgnoreCase))
					{
						Console.WriteLine("\r\nNot overwriting existing file.\r\nExiting ExcelDnaPack.");
						return 1;
					}
				}

				try
				{
					File.Delete(xllOutputPath);
				}
				catch
				{
					Console.Error.Write("ERROR: Existing output .xll file " + xllOutputPath + "could not be deleted. (Perhaps loaded in Excel?)\r\n\r\nExiting ExcelDnaPack.");
					return 1;
				}
			}

            string outputDirectory = Path.GetDirectoryName(xllOutputPath);
            if (outputDirectory == String.Empty)
            {
                outputDirectory = ".";  // https://github.com/Excel-DNA/ExcelDna/issues/7
            }

            if (!Directory.Exists(outputDirectory))
            {
				try
				{
                    Directory.CreateDirectory(outputDirectory);
                }
                catch (Exception ex)
                {
                    Console.Error.Write("ERROR: Output directory " + outputDirectory + "could not be created. Error: " + ex.Message + "\r\n\r\nExiting ExcelDnaPack.");
                    return 1;
                }
            }

			// Find ExcelDna.xll to use.
            // First try <MyAddin>.xll
            if (!File.Exists(xllInputPath))
            {
                // CONSIDER: Maybe the next two (old) search locations should be deprecated?
                // Then try one called ExcelDna.xll next to the .dna file
                xllInputPath = Path.Combine(dnaDirectory, "ExcelDna.xll");
                if (!File.Exists(xllInputPath))
                {
                    // Then try one called ExcelDna.xll next to the ExcelDnaPack.exe
                    xllInputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelDna.xll");
                    if (!File.Exists(xllInputPath))
                    {
                        Console.Error.WriteLine("ERROR: Base add-in not found.\r\n\r\n" + usageInfo);
                        return 1;
                    }
                }
            }
			Console.WriteLine("Using base add-in " + xllInputPath);

			File.Copy(xllInputPath, xllOutputPath, false);
			ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(xllOutputPath);
			if (File.Exists(configPath))
			{
				ru.AddFile(File.ReadAllBytes(configPath), "__MAIN__",ResourceHelper.TypeName.CONFIG, false, multithreading);  // Name here must exactly match name in ExcelDnaLoad.cpp.
			}
			byte[] dnaBytes = File.ReadAllBytes(dnaPath);
			byte[] dnaContentForPacking = PackDnaLibrary(dnaBytes, dnaDirectory, ru, compress, multithreading);
			ru.AddFile(dnaContentForPacking, "__MAIN__",ResourceHelper.TypeName.DNA, false, multithreading); // Name here must exactly match name in DnaLibrary.Initialize.
			ru.EndUpdate();
			Console.WriteLine("Completed Packing {0}.", xllOutputPath);
#if DEBUG
			Console.WriteLine("Press any key to exit.");
			Console.ReadKey();
#endif
            // All OK - set process exit code to 'Success'
            return 0;
        }

		static int lastPackIndex = 0;

		static byte[] PackDnaLibrary(byte[] dnaContent, string dnaDirectory, ResourceHelper.ResourceUpdater ru, bool compress, bool multithreading)
        {
            string errorMessage;
			DnaLibrary dna = DnaLibrary.LoadFrom(dnaContent, dnaDirectory);
            if (dna == null)
            {
                // TODO: Better error handling here.
                errorMessage = "ERROR: .dna file could not be loaded. Possibly malformed xml content? ABORTING.";
                throw new InvalidOperationException(errorMessage);
            }
			if (dna.ExternalLibraries != null)
			{
			    bool copiedVersion = false;
				foreach (ExternalLibrary ext in dna.ExternalLibraries)
				{
					var path = dna.ResolvePath(ext.Path);
					if (!File.Exists(path))
					{
						errorMessage = string.Format("!!! ERROR: ExternalLibrary `{0}` not found. ABORTING.", ext.Path);
                        throw new InvalidOperationException(errorMessage);
					}

					if (ext.Pack)
					{
                        Console.WriteLine("  ~~> ExternalLibrary path {0} resolved to {1}.", ext.Path, path);
						if (Path.GetExtension(path).Equals(".DNA", StringComparison.OrdinalIgnoreCase))
						{
							string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + ".DNA";
							byte[] dnaContentForPacking = PackDnaLibrary(File.ReadAllBytes(path), Path.GetDirectoryName(path), ru, compress, multithreading);
							ru.AddFile(dnaContentForPacking, name, ResourceHelper.TypeName.DNA, compress, multithreading);
							ext.Path = "packed:" + name;
						}
						else
						{
							string packedName = ru.AddAssembly(path, compress, multithreading, ext.IncludePdb);
							if (packedName != null)
							{
								ext.Path = "packed:" + packedName;
							}
						}
                        if (ext.ComServer == true)
                        {
                            // Check for a TypeLib to pack
                            //string tlbPath = dna.ResolvePath(ext.TypeLibPath);
                            string resolvedTypeLibPath = null;
                            if (!string.IsNullOrEmpty(ext.TypeLibPath))
                            {
                                resolvedTypeLibPath = dna.ResolvePath(ext.TypeLibPath); // null is unresolved
                                if (resolvedTypeLibPath == null)
                                {
                                    // Try relative to .dll
                                    resolvedTypeLibPath = DnaLibrary.ResolvePath(ext.TypeLibPath, System.IO.Path.GetDirectoryName(path) ); // null is unresolved
                                    if (resolvedTypeLibPath == null)
                                    {
                                        errorMessage = string.Format("!!! ERROR: ExternalLibrary TypeLib path {0} could not be resolved.", ext.TypeLibPath);
                                        throw new InvalidOperationException(errorMessage);
                                    }
                                }
                            }
                            else
                            {
                                // Check for .tlb
                                string tlbCheck = System.IO.Path.ChangeExtension(path, "tlb");
                                if (System.IO.File.Exists(tlbCheck))
                                {
                                    resolvedTypeLibPath = tlbCheck;
                                }
                            }
                            if (resolvedTypeLibPath != null)
                            {
                                Console.WriteLine("  ~~> ExternalLibrary typelib path {0} resolved to {1}.", ext.TypeLibPath, resolvedTypeLibPath);
                                int packedIndex = ru.AddTypeLib(File.ReadAllBytes(resolvedTypeLibPath));
                                ext.TypeLibPath = "packed:" + packedIndex.ToString();
                            }
                        }
					}
				    if (ext.UseVersionAsOutputVersion)
				    {
				        if (copiedVersion)
				        {
				            Console.WriteLine("  ~~> Assembly version already copied from previous ExternalLibrary; ignoring 'UseVersionAsOutputVersion' attribute.");
				            continue;
				        }
				        try
				        {
				            ru.CopyFileVersion(path);
				            copiedVersion = true;
				        }
				        catch (Exception e)
				        {
                            errorMessage = string.Format("  ~~> ERROR: Error copying version to output version: {0}", e.Message);
                            throw new InvalidOperationException(errorMessage);
				        }
				    }
				}
			}
			// Collect the list of all the references.
			List<Reference> refs = new List<Reference>();
			foreach (Project proj in dna.GetProjects())
			{
				if (proj.References != null)
				{
					refs.AddRange(proj.References);
				}
			}
			// Fix-up if Reference is not part of a project, but just used to add an assembly for packing.
			foreach (Reference rf in dna.References)
			{
				if (!refs.Contains(rf))
					refs.Add(rf);
			}

			// Expand asterisk in filename of reference path, e.g. "./*.dll"
			List<Reference> expandedReferences = new List<Reference>();
			for (int i = refs.Count - 1; i >= 0; i--)
			{
				string path = refs[i].Path; 
				if (path != null && path.Contains("*"))
				{
					var files = Directory.GetFiles(Path.GetDirectoryName(path), Path.GetFileName(path));
					foreach (var file in files)
					{
						expandedReferences.Add(new Reference(Path.GetFullPath(file)) {Pack = true});
					}
					refs.RemoveAt(i);
				}
			}
			refs.AddRange(expandedReferences);

			// Now pack the references
			foreach (Reference rf in refs)
			{
				if (rf.Pack)
				{
					string path = null;
					if (rf.Path != null)
					{
						if (rf.Path.StartsWith("packed:"))
						{
							continue;
						}

                        path = dna.ResolvePath(rf.Path);
                        Console.WriteLine("  ~~> Assembly path {0} resolved to {1}.", rf.Path, path);
                    }
					if (path == null && rf.Name != null)
					{
						// Try Load as as last resort (and opportunity to load by FullName)
						try
                        {
                            #pragma warning disable 0618
                            Assembly ass = Assembly.LoadWithPartialName(rf.Name);
                            #pragma warning restore 0618
                            if (ass != null)
							{
								path = ass.Location;
								Console.WriteLine("  ~~> Assembly {0} 'Load'ed from location {1}.", rf.Name, path);
							}
						}
						catch (Exception e)
						{
							Console.WriteLine("  ~~> Assembly {0} not 'Load'ed. Exception: {1}", rf.Name, e);
						}
					}
					if (path == null)
					{
                        errorMessage = string.Format("  ~~> ERROR: Reference with Path: {0} and Name: {1} NOT FOUND.", rf.Path, rf.Name);
                        throw new InvalidOperationException(errorMessage);
					}
					
					// It worked!
					string packedName = ru.AddAssembly(path, compress, multithreading, rf.IncludePdb);
					if (packedName != null)
					{
						rf.Path = "packed:" + packedName;
					}
				}
			}
            foreach (Image image in dna.Images)
            {
                if (image.Pack)
                {
                    string path = dna.ResolvePath(image.Path);
                    if (path == null)
                    {
                        errorMessage = string.Format("  ~~> ERROR: Image path {0} NOT RESOLVED.", image.Path);
                        throw new InvalidOperationException(errorMessage);
                    }
                    string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                    byte[] imageBytes = File.ReadAllBytes(path);
                    ru.AddFile(imageBytes, name, ResourceHelper.TypeName.IMAGE, compress, multithreading);
                    image.Path = "packed:" + name;
                }
            }
            foreach (Project project in dna.Projects)
            {
                foreach (SourceItem source in project.SourceItems)
                {
                    if (source.Pack && !string.IsNullOrEmpty(source.Path))
                    {
                        string path = dna.ResolvePath(source.Path);
                        if (path == null)
                        {
                            errorMessage = string.Format("  ~~> ERROR: Source path {0} NOT RESOLVED.", source.Path);
                            throw new InvalidOperationException(errorMessage);
                        }
                        string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                        byte[] sourceBytes = File.ReadAllBytes(path);
                        ru.AddFile(sourceBytes, name, ResourceHelper.TypeName.SOURCE, compress, multithreading);
                        source.Path = "packed:" + name;
                    }
                }
            }
		    return DnaLibrary.Save(dna);
		}


        static void PackXllBuild(string xllFullPath, bool includePdb)
        {
            ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(xllFullPath);

            var xllDir = Path.GetDirectoryName(xllFullPath);
            ru.AddAssembly(Path.Combine(xllDir, "ExcelDna.ManagedHost.dll"), compress: false, multithreading: false, includePdb);
            ru.AddAssembly(Path.Combine(xllDir, "ExcelDna.Loader.dll"), compress: true, multithreading: false, includePdb);
            ru.AddAssembly(Path.Combine(xllDir, "ExcelDna.Integration.dll"), compress: true, multithreading: false, includePdb);
            ru.EndUpdate();
        }
	}

}
