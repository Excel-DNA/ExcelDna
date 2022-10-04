using System;
using System.Collections.Generic;
using System.IO;
using ExcelDna.Integration;
using System.Reflection;

namespace ExcelDna.PackedResources
{
    internal class ExcelDnaPack
    {
        public static int Pack(string dnaPath, string xllOutputPathParam, bool compress, bool multithreading, bool overwrite, string usageInfo, List<string> filesToPublish, bool useManagedResourceResolver)
        {
            string dnaDirectory = Path.GetDirectoryName(dnaPath);
            string dnaFilePrefix = Path.GetFileNameWithoutExtension(dnaPath);
            string configPath = Path.ChangeExtension(dnaPath, ".xll.config");
            string xllInputPath = Path.Combine(dnaDirectory, dnaFilePrefix + ".xll");
            string xllOutputPath = Path.Combine(dnaDirectory, dnaFilePrefix + "-packed.xll");
            if (xllOutputPathParam != null)
                xllOutputPath = xllOutputPathParam;

            if (!File.Exists(dnaPath))
            {
                Console.Error.Write("ERROR: Add-in .dna file " + dnaPath + " not found.\r\n\r\n" + usageInfo);
                return 1;
            }

            if (filesToPublish == null && File.Exists(xllOutputPath))
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

            if (filesToPublish == null && !Directory.Exists(outputDirectory))
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

            if (useManagedResourceResolver)
                Console.WriteLine("Using managed resource packing.");

            ResourceHelper.ResourceUpdater ru = null;
            if (filesToPublish == null)
            {
                File.Copy(xllInputPath, xllOutputPath, false);
                ru = new ResourceHelper.ResourceUpdater(Path.Combine(Directory.GetCurrentDirectory(), xllOutputPath), useManagedResourceResolver);
            }
            else
            {
                filesToPublish.Add(xllInputPath);
            }
            if (File.Exists(configPath))
            {
                if (filesToPublish == null)
                    ru.AddFile(File.ReadAllBytes(configPath), "__MAIN__", ResourceHelper.TypeName.CONFIG, false, multithreading);  // Name here must exactly match name in ExcelDnaLoad.cpp.
                else
                    filesToPublish.Add(configPath);
            }
            byte[] dnaBytes = File.ReadAllBytes(dnaPath);
            byte[] dnaContentForPacking = PackDnaLibrary(dnaBytes, dnaDirectory, ru, compress, multithreading, filesToPublish);
            if (filesToPublish == null)
            {
                ru.AddFile(dnaContentForPacking, "__MAIN__", ResourceHelper.TypeName.DNA, false, multithreading); // Name here must exactly match name in DnaLibrary.Initialize.
                ru.EndUpdate();
            }
            else
            {
                filesToPublish.Add(dnaPath);
            }
            Console.WriteLine("Completed Packing {0}.", xllOutputPath);

            // All OK - set process exit code to 'Success'
            return 0;
        }

        static private byte[] PackDnaLibrary(byte[] dnaContent, string dnaDirectory, ResourceHelper.ResourceUpdater ru, bool compress, bool multithreading, List<string> filesToPublish)
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
                            byte[] dnaContentForPacking = PackDnaLibrary(File.ReadAllBytes(path), Path.GetDirectoryName(path), ru, compress, multithreading, filesToPublish);
                            if (filesToPublish == null)
                            {
                                ru.AddFile(dnaContentForPacking, name, ResourceHelper.TypeName.DNA, compress, multithreading);
                                ext.Path = "packed:" + name;
                            }
                            else
                            {
                                filesToPublish.Add(path);
                            }
                        }
                        else
                        {
                            if (filesToPublish == null)
                            {
                                string packedName = ru.AddAssembly(path, compress, multithreading, ext.IncludePdb);
                                if (packedName != null)
                                {
                                    ext.Path = "packed:" + packedName;
                                }
                            }
                            else
                            {
                                filesToPublish.Add(path);
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
                                    resolvedTypeLibPath = DnaLibrary.ResolvePath(ext.TypeLibPath, System.IO.Path.GetDirectoryName(path)); // null is unresolved
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
                                if (filesToPublish == null)
                                {
                                    Console.WriteLine("  ~~> ExternalLibrary typelib path {0} resolved to {1}.", ext.TypeLibPath, resolvedTypeLibPath);
                                    int packedIndex = ru.AddTypeLib(File.ReadAllBytes(resolvedTypeLibPath));
                                    ext.TypeLibPath = "packed:" + packedIndex.ToString();
                                }
                                else
                                {
                                    filesToPublish.Add(resolvedTypeLibPath);
                                }
                            }
                        }
                    }
                    if (filesToPublish == null && ext.UseVersionAsOutputVersion)
                    {
                        if (copiedVersion)
                        {
                            Console.WriteLine("  ~~> Assembly version already copied from previous ExternalLibrary; ignoring 'UseVersionAsOutputVersion' attribute.");
                            continue;
                        }
                        try
                        {
                            ru.CopyFileVersion(Path.Combine(Directory.GetCurrentDirectory(), path));
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
                        expandedReferences.Add(new Reference(Path.GetFullPath(file)) { Pack = true });
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
                    if (filesToPublish == null)
                    {
                        string packedName = ru.AddAssembly(path, compress, multithreading, rf.IncludePdb);
                        if (packedName != null)
                        {
                            rf.Path = "packed:" + packedName;
                        }
                    }
                    else
                    {
                        filesToPublish.Add(path);
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
                    if (filesToPublish == null)
                    {
                        string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                        byte[] imageBytes = File.ReadAllBytes(path);
                        ru.AddFile(imageBytes, name, ResourceHelper.TypeName.IMAGE, compress, multithreading);
                        image.Path = "packed:" + name;
                    }
                    else
                    {
                        filesToPublish.Add(path);
                    }
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
                        if (filesToPublish == null)
                        {
                            string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                            byte[] sourceBytes = File.ReadAllBytes(path);
                            ru.AddFile(sourceBytes, name, ResourceHelper.TypeName.SOURCE, compress, multithreading);
                            source.Path = "packed:" + name;
                        }
                        else
                        {
                            filesToPublish.Add(path);
                        }
                    }
                }
            }
            return DnaLibrary.Save(dna);
        }

        static private int lastPackIndex = 0;
    }
}
