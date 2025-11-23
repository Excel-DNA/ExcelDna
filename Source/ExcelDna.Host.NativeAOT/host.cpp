//  Copyright (c) Govert van Drimmelen. All rights reserved.

#include "host.h"
#include "exports.h"
#include "path.h"
#include "TempDir.h"
#include "utils.h"
#include "dnainfo.h"

#include <filesystem>

TempDir tempDir(L"ExcelDna.Host.NativeAOT");

int load_and_run(const std::wstring& basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll)
{
	std::wstring hostFileName;
	{
		int r = TryLoadPropertyFromResource(hModuleXll, L"__MAIN___ORIGINAL_DLL_FILE_NAME", hostFileName);
		if (r != EXIT_SUCCESS)
		{
			hostFileName = GetFileName(GetAddInFullPath());
			RenameExtension(hostFileName, L".dll");
		}
	}
	std::wstring hostFile(PathCombine(GetDirectory(GetAddInFullPath()), hostFileName));

	if (!std::filesystem::exists(hostFile))
	{
		hostFile = PathCombine(tempDir.GetPath(), hostFileName);
		if (!std::filesystem::exists(hostFile))
		{
			int r = WriteResourceToFile(hModuleXll, L"__MAIN__", L"NATIVE_ASSEMBLY", hostFile);
			if (r != EXIT_SUCCESS)
				return r;
		}
	}

	if (FindResource(hModuleXll, L"__MAIN__", L"PDB") != NULL)
	{
		std::wstring pdbFileName;
		{
			int r = LoadPropertyFromResource(hModuleXll, L"__MAIN___ORIGINAL_PDB_FILE_NAME", pdbFileName);
			if (r != EXIT_SUCCESS)
				return r;
		}
		std::wstring pdbFile(PathCombine(GetDirectory(hostFile), pdbFileName));

		if (!std::filesystem::exists(pdbFile))
		{
			int r = WriteResourceToFile(hModuleXll, L"__MAIN__", L"PDB", pdbFile);
			if (r != EXIT_SUCCESS)
				return r;
		}
	}

	HINSTANCE handle = LoadLibrary(hostFile.c_str());

	if (handle == NULL)
	{
		ShowHostError(L"Loading " + hostFile + L" library failed.");
		return EXIT_FAILURE;
	}

	typedef short(__stdcall* xladdin_initialize_native_fn)(void* xlAddInExportInfo, void* hModuleXLL, void* pPathXLL, BYTE disableAssemblyContextUnload, void* pTempDirPath);

	xladdin_initialize_native_fn init = (xladdin_initialize_native_fn)GetProcAddress(handle, "Initialize");
	if (init == NULL)
	{
		ShowHostError(L"GetProcAddress Initialize failed.");
		return EXIT_FAILURE;
	}

	std::wstring tempDirPath = tempDir.GetPath();
	short res = init(pExportInfo, hModuleXll, (void*)pathXll, false, (void*)tempDirPath.c_str());

	return res == 0 ? EXIT_FAILURE : EXIT_SUCCESS;
}
