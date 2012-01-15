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

  dnaPath      The path to the primary .dna file for the ExcelDna add-in.
  /Y           If the output .xll exists, overwrite without prompting.
  /O outPath   Output path - default is <dnaPath>-packed.xll.

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
		
		static void Main(string[] args)
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
				return;
			}

			string dnaPath = args[0];
			string dnaDirectory = Path.GetDirectoryName(dnaPath);
//			string dnaFileName = Path.GetFileName(dnaPath);
			string dnaFilePrefix = Path.GetFileNameWithoutExtension(dnaPath);
			string configPath = Path.ChangeExtension(dnaPath, ".xll.config");
            string xllInputPath = Path.Combine(dnaDirectory, dnaFilePrefix + ".xll");
			string xllOutputPath = Path.Combine(dnaDirectory, dnaFilePrefix + "-packed.xll");
            bool overwrite = false;

			if (!File.Exists(dnaPath))
			{
				Console.Write("Add-in .dna file " + dnaPath + " not found.\r\n\r\n" + usageInfo);
				return;
			}

            // TODO: Replace with an args-parsing routine.
            if (args.Length > 1)
            {
                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].ToUpper() == "/O")
                    {
                        if (i >= args.Length - 1)
                        {
                            // Too few args.
                            Console.Write("Invalid command-line arguments.\r\n\r\n" + usageInfo);
                            return;
                        }
                        xllOutputPath = args[i + 1];
                    }
                    else if (args[i].ToUpper() == "/Y")
                    {
                        overwrite = true;
                    }
                }
            }

			if (File.Exists(xllOutputPath))
			{
				if (overwrite == false)
				{
					Console.Write("Output .xll file " + xllOutputPath + " already exists. Overwrite? [Y/N] ");
					string response = Console.ReadLine();
					if (response.ToUpper() != "Y")
					{
						Console.WriteLine("\r\nNot overwriting existing file.\r\nExiting ExcelDnaPack.");
						return;
					}
				}

				try
				{
					File.Delete(xllOutputPath);
				}
				catch
				{
					Console.Write("Existing output .xll file " + xllOutputPath + "could not be deleted. (Perhaps loaded in Excel?)\r\n\r\nExiting ExcelDnaPack.");
					return;
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
                        Console.WriteLine("Base add-in not found.\r\n\r\n" + usageInfo);
                        return;
                    }
                }
            }
			Console.WriteLine("Using base add-in " + xllInputPath);

			File.Copy(xllInputPath, xllOutputPath, false);
			ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(xllOutputPath);
			// Take out Integration assembly - to be replaced by a compressed copy.
            // CONSIDER: Maybe use the ExcelDna.Integration that is inside the <MyAddin>.xll
			ru.RemoveResource("ASSEMBLY", "EXCELDNA.INTEGRATION");
            string integrationPath = DnaLibrary.ResolvePath("ExcelDna.Integration.dll", dnaDirectory);
			string packedName = null;
			if (integrationPath != null)
			{
				packedName = ru.AddAssembly(integrationPath);
			}
			if (packedName == null)
			{
				Console.WriteLine("ExcelDna.Integration assembly could not be packed. Aborting.");
				ru.EndUpdate();
				File.Delete(xllOutputPath);
				return;
			}
			if (File.Exists(configPath))
			{
				ru.AddConfigFile(File.ReadAllBytes(configPath), "__MAIN__");  // Name here must exactly match name in ExcelDnaLoad.cpp.
			}
			byte[] dnaBytes = File.ReadAllBytes(dnaPath);
			byte[] dnaContentForPacking = PackDnaLibrary(dnaBytes, dnaDirectory, ru);
			ru.AddDnaFileUncompressed(dnaContentForPacking, "__MAIN__"); // Name here must exactly match name in DnaLibrary.Initialize.
			ru.EndUpdate();
			Console.WriteLine("Completed Packing {0}.", xllOutputPath);
#if DEBUG
			Console.WriteLine("Press any key to exit.");
			Console.ReadKey();
#endif
		}

		static int lastPackIndex = 0;

		static byte[] PackDnaLibrary(byte[] dnaContent, string dnaDirectory, ResourceHelper.ResourceUpdater ru)
		{
			DnaLibrary dna = DnaLibrary.LoadFrom(dnaContent, dnaDirectory);
            if (dna == null)
            {
                // TODO: Better error handling here.
                Console.WriteLine(".dna file could not be loaded. Possibly malformed xml content? Aborting.");
                Environment.Exit(1);
            }
			if (dna.ExternalLibraries != null)
			{
				foreach (ExternalLibrary ext in dna.ExternalLibraries)
				{
					if (ext.Pack)
					{
						string path = dna.ResolvePath(ext.Path);
                        Console.WriteLine("  ~~> ExternalLibrary path {0} resolved to {1}.", ext.Path, path);
						if (Path.GetExtension(path).ToUpperInvariant() == ".DNA")
						{
							string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + ".DNA";
							byte[] dnaContentForPacking = PackDnaLibrary(File.ReadAllBytes(path), Path.GetDirectoryName(path), ru);
							ru.AddDnaFile(dnaContentForPacking, name);
							ext.Path = "packed:" + name;
						}
						else
						{
							string packedName = ru.AddAssembly(path);
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
                                        Console.WriteLine("!!! ExternalLibrary TypeLib path {0} could not be resolved.", ext.TypeLibPath);
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
							break;
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
						Console.WriteLine("  ~~> Reference with Path: {0} and Name: {1} not found.", rf.Path, rf.Name);
						break;
					}
					
					// It worked!
					string packedName = ru.AddAssembly(path);
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
                        Console.WriteLine("  ~~> Image path {0} not resolved.", image.Path);
                        break;
                    }
                    string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                    byte[] imageBytes = File.ReadAllBytes(path);
                    ru.AddImage(imageBytes, name);
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
                            Console.WriteLine("  ~~> Source path {0} not resolved.", source.Path);
                            break;
                        }
                        string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                        byte[] sourceBytes = Encoding.UTF8.GetBytes(File.ReadAllText(path));
                        ru.AddSource(sourceBytes, name);
                        source.Path = "packed:" + name;
                    }
                }
            }
		    return DnaLibrary.Save(dna);
		}

	}

}
