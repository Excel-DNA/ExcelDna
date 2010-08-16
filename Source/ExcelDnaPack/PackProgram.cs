using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;

using System.IO;
using System.Diagnostics;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace ExcelDnaPack
{
	class PackProgram
	{
		static string usageInfo = 
@"ExcelDnaPack Usage
------------------
ExcelDnaPack is a command-line utility to pack ExcelDna add-ins into a single .xll file.

Usage: ExcelDnaPack.exe dnaPath [/Y]

  dnaPath      The path to the primary .dna file for the ExcelDna add-in.
  /Y           If output .xll exists, overwrite without prompting.

Example: ExcelDnaPack.exe MyAddins\FirstAddin.dna
		 The packed add-in file will be created as MyAddins\FirstAddin.xll.

The template add-in host file (always called ExcelDna.xll) and helper 
assembly (ExcelDna.Integration.dll) are searched for 
  1. in the same directory as the .dna file, and if not found there, 
  2. in the same directory as the ExcelDnaPack.exe file.

ExcelDnaPack will also pack the configuration file FirstAddin.xll.config if it is 
found next to FirstAddin.dna.
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

			string xllOutputPath = Path.Combine(dnaDirectory, dnaFilePrefix + "-packed.xll");

			if (!File.Exists(dnaPath))
			{
				Console.Write("Add-in .dna file " + dnaPath + " not found.\r\n\r\n" + usageInfo);
				return;
			}
			if (File.Exists(xllOutputPath))
			{
				if (args.Length == 1)
				{
					Console.Write("Output .xll file " + xllOutputPath + " already exists. Overwrite? [Y/N] ");
					string response = Console.ReadLine();
					if (response.ToUpper() != "Y")
					{
						Console.WriteLine("\r\nNot overwriting existing file.\r\nExiting ExcelDnaPack.");
						return;
					}
				}
				else if (args[1].ToUpper() != "/Y")
				{
					Console.Write("Invalid command-line arguments.\r\n\r\n" + usageInfo);
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
			string xllInputPath;
			xllInputPath = Path.Combine(dnaDirectory, "ExcelDna.xll");
			if (!File.Exists(xllInputPath))
			{
				xllInputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelDna.xll");
				if (!File.Exists(xllInputPath))
				{
					Console.WriteLine("Base add-in not found.\r\n\r\n" + usageInfo);
					return;
				}
			}
			Console.WriteLine("Using base add-in " + xllInputPath);

			File.Copy(xllInputPath, xllOutputPath, false);
			ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(xllOutputPath);
			// Take out Integration assembly - to be replaced by a compressed copy.
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
			if (dna.ExternalLibraries != null)
			{
				foreach (var ext in dna.ExternalLibraries)
				{
					if (ext.Pack)
					{
						string path = dna.ResolvePath(ext.Path);
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
					}
				}
			}
			// Collect the list of all the references.
			List<Reference> refs = new List<Reference>();
			foreach (var proj in dna.GetProjects())
			{
				if (proj.References != null)
				{
					refs.AddRange(proj.References);
				}
			}
			// Fix-up if Reference is not part of a project, but just used to add an assembly for packing.
			foreach (var rf in dna.References)
			{
				if (!refs.Contains(rf))
					refs.Add(rf);
			}
			// Now pack the references
			foreach (var rf in refs)
			{
				if (rf.Pack)
				{
					string path = null;
					if (rf.AssemblyPath != null)
					{
						if (rf.AssemblyPath.StartsWith("packed:"))
						{
							break;
						}

                        path = dna.ResolvePath(rf.AssemblyPath);

						if (path == null)
						{
							Console.WriteLine("  ~~> Assembly path {0} not resolved.", rf.AssemblyPath);
							// Try Load as as last resort (and opportunity to load by FullName)
							try
							{
								Assembly ass = Assembly.LoadWithPartialName(rf.AssemblyPath);
								if (ass != null)
								{
									path = ass.Location;
									Console.WriteLine("  ~~> Assembly {0} 'Load'ed from location {1}.", rf.AssemblyPath, path);
								}
							}
							catch (Exception e)
							{
								Console.WriteLine("  ~~> Assembly {0} not 'Load'ed. Exception: {1}", rf.AssemblyPath, e);
							}
						}
					}
					if (path == null && rf.Name != null)
					{
							if (path == null)
							{
								// Try Load as as last resort (and opportunity to load by FullName)
								try
								{
									Assembly ass = Assembly.LoadWithPartialName(rf.Name);
									if (ass != null)
									{
										path = ass.Location;
										Console.WriteLine("  ~~> Assembly {0} loaded via partial name from {1}.", rf.Name, path);
									}
								}
								catch (Exception e)
								{
									Console.WriteLine("  ~~> Assembly {0} not 'Load'ed. Exception: {1}", rf.AssemblyPath, e);
								}
							}
						}
					if (path == null)
					{
						Console.WriteLine("  ~~> Reference with AssemblyPath: {0}, Name: {1} not found.", rf.AssemblyPath, rf.Name);
						break;
					}
					
					// It worked!
					string packedName = ru.AddAssembly(path);
					if (packedName != null)
					{
						rf.AssemblyPath = "packed:" + packedName;
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
                    string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant(); ;
                    byte[] imageBytes = File.ReadAllBytes(path);
                    ru.AddImage(imageBytes, name);
                    image.Path = "packed:" + name;
                }
            }

			return DnaLibrary.Save(dna);
		}

	}

}
