//  Copyright (c) Govert van Drimmelen. All rights reserved.

// xllhost.cpp : Defines the entry point for the XLL host
#include "host.h"
#include "exports.h"

// Include the official nethost API and indicate
// consumption should be as a static library.
#define NETHOST_USE_AS_STATIC
#include <nethost.h>

#include <coreclr_delegates.h>
#include <hostfxr.h>
#include <error_codes.h>

#include <cassert>
#include <iostream>
#include <string>
#include <filesystem>

#include "TempDir.h"
#include "utils.h"
#include "dnainfo.h"

using string_t = std::basic_string<char_t>;

// Globals to hold hostfxr exports
hostfxr_initialize_for_runtime_config_fn init_fptr;
hostfxr_get_runtime_delegate_fn get_delegate_fptr;
hostfxr_get_runtime_property_value_fn get_property_fptr;
hostfxr_close_fn close_fptr;

// Forward declarations
bool load_hostfxr(int& rc);
load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly();

// Provide a callback for any catastrophic failures.
// The provided callback will be the last call prior to a rude-abort of the process.
// void __stdcall set_failure_callback(failure_fn cb) {}

// Preload the runtime.
// The runtime is lazily loaded whenever the first export is called. This function
// preloads the runtime independent of calling any export and avoids the startup
// cost associated with calling an export for the first time.
void __stdcall preload_runtime(void) {}

TempDir tempDir(L"ExcelDna.Host");

// TODO: Might return the fn*
int load_runtime_and_run(const std::wstring& basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll)
{
	//
	// STEP 1: Load HostFxr and get exported hosting functions
	//
	int rc = 0;
	if (!load_hostfxr(rc))
	{
		std::wstring msg;
		if (rc == CoreHostLibMissingFailure)
		{
#if _WIN64
			std::wstring bitness = L"x64";
#else
			std::wstring bitness = L"x86";
#endif
			std::wstring runtime = L".NET Desktop Runtime 6.0 " + bitness;
			msg = std::format(L"{0} is not installed, corrupted or incomplete.\n\nYou can download {0} from https://dotnet.microsoft.com/en-us/download/dotnet/6.0", runtime);
		}
		else if (rc != 0)
		{
			msg = std::format(L"Failure: load_hostfxr().\n\nError code: {}.", rc);
		}
		else
		{
			msg = L"Failure: load_hostfxr().";
		}
		ShowHostError(msg);

		return EXIT_FAILURE;
	}

	//
	// STEP 2: Initialize and start the .NET Core runtime
	//
	load_assembly_and_get_function_pointer_fn load_assembly_and_get_function_pointer = get_dotnet_load_assembly();
	if (load_assembly_and_get_function_pointer == nullptr)
		return EXIT_FAILURE;

	//
	// STEP 3: Copy managed assembly from resources to some temp file if needed
	//
	std::wstring hostFile = PathCombine(basePath, L"ExcelDna.ManagedHost.dll");
	if (!std::filesystem::exists(hostFile))
	{
		hostFile = PathCombine(tempDir.GetPath(), L"ExcelDna.ManagedHost.dll");
		if (!std::filesystem::exists(hostFile))
		{
			HRSRC hResManagedHost = FindResource(hModuleXll, L"EXCELDNA.MANAGEDHOST", L"ASSEMBLY");
			if (hResManagedHost == NULL)
			{
				ShowHostError(L"Failure to find resource EXCELDNA.MANAGEDHOST");
				return EXIT_FAILURE;
			}

			HGLOBAL hManagedHost = LoadResource(hModuleXll, hResManagedHost);
			if (hManagedHost == NULL)
			{
				ShowHostError(L"Failure to load resource EXCELDNA.MANAGEDHOST");
				return EXIT_FAILURE;
			}

			void* buf = LockResource(hManagedHost);
			if (buf == NULL)
			{
				ShowHostError(L"Failure to lock resource EXCELDNA.MANAGEDHOST");
				return EXIT_FAILURE;
			}

			DWORD resSize = SizeofResource(hModuleXll, hResManagedHost);
			SafeByteArray safeBytes(buf, resSize);
			byte* pData;
			int nSize = safeBytes.AccessData(&pData);

			HRESULT hr = WriteAllBytes(hostFile, pData, nSize);
			if (FAILED(hr))
			{
				std::wstringstream stream;
				stream << "Saving EXCELDNA.MANAGEDHOST failed: " << std::hex << std::showbase << hr;
				ShowHostError(stream.str());
				return EXIT_FAILURE;
			}
		}
	}

	//
	// STEP 4: Load managed assembly and get function pointer to a managed method
	//
	const string_t dotnetlib_path = hostFile;
	const char_t* dotnet_type = L"ExcelDna.ManagedHost.AddInInitialize, ExcelDna.ManagedHost";
	const char_t* dotnet_type_method = L"Initialize";

	// Function pointer to managed delegate with non-default signature
	typedef short (CORECLR_DELEGATE_CALLTYPE* xladdin_initialize_fn)(void* xlAddInExportInfo, void* hModuleXLL, void* pPathXLL, BYTE disableAssemblyContextUnload, void* pTempDirPath);
	xladdin_initialize_fn init = nullptr;
	rc = load_assembly_and_get_function_pointer(
		dotnetlib_path.c_str(),
		dotnet_type,
		dotnet_type_method,
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		(void**)&init);
	assert(rc == 0 && init != nullptr && "Failure: load_assembly_and_get_function_pointer()");

	bool disableAssemblyContextUnload;
	HRESULT hr = GetDisableAssemblyContextUnload(disableAssemblyContextUnload);
	if (FAILED(hr))
		disableAssemblyContextUnload = false;

	std::wstring tempDirPath = tempDir.GetPath();
	short res = init(pExportInfo, hModuleXll, (void*)pathXll, disableAssemblyContextUnload, (void*)tempDirPath.c_str());

	return res == 0 ? EXIT_FAILURE : EXIT_SUCCESS;
}


/********************************************************************************************
 * Function used to load and activate .NET Core
 ********************************************************************************************/

void* load_library(const char_t* path)
{
	HMODULE h = ::LoadLibraryW(path);
	assert(h != nullptr);
	return (void*)h;
}
void* get_export(void* h, const char* name)
{
	void* f = ::GetProcAddress((HMODULE)h, name);
	assert(f != nullptr);
	return f;
}

// Using the nethost library, discover the location of hostfxr and get exports
bool load_hostfxr(int& rc)
{
	// Pre-allocate a large buffer for the path to hostfxr
	char_t buffer[MAX_PATH];
	size_t buffer_size = sizeof(buffer) / sizeof(char_t);
	rc = get_hostfxr_path(buffer, &buffer_size, nullptr);
	if (rc != 0)
		return false;

	// Load hostfxr and get desired exports
	void* lib = load_library(buffer);
	init_fptr = (hostfxr_initialize_for_runtime_config_fn)get_export(lib, "hostfxr_initialize_for_runtime_config");
	get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)get_export(lib, "hostfxr_get_runtime_delegate");
	get_property_fptr = (hostfxr_get_runtime_property_value_fn)get_export(lib, "hostfxr_get_runtime_property_value");
	close_fptr = (hostfxr_close_fn)get_export(lib, "hostfxr_close");

	return (init_fptr && get_delegate_fptr && close_fptr);
}

std::wstring get_runtime_property(const hostfxr_handle host_context_handle, const std::wstring& name)
{
	const wchar_t* value = NULL;
	int rc = get_property_fptr(host_context_handle, name.c_str(), &value);
	if (STATUS_CODE_SUCCEEDED(rc) && value != NULL)
	{
		return value;
	}

	return L"";
}

std::wstring get_loaded_runtime_version()
{
	return GetDirectoryName(get_runtime_property(NULL, L"FX_DEPS_FILE"));
}

// Load and initialize .NET Core and get desired function pointer for scenario
load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly()
{
	std::string configText = R"({
  "runtimeOptions": {
    "tfm": "net6.0",
    "framework": {
      "name": "Microsoft.WindowsDesktop.App",
      "version": "6.0.0"
    }
  }
})";
	std::wstring configFile = PathCombine(tempDir.GetPath(), L"ExcelDna.Host.runtimeconfig.json");
	HRESULT hr = WriteAllBytes(configFile, (void*)configText.c_str(), static_cast<DWORD>(configText.length()));
	if (FAILED(hr))
	{
		std::wstringstream stream;
		stream << "Saving ExcelDna.Host.runtimeconfig.json failed: " << std::hex << std::showbase << hr;
		ShowHostError(stream.str());
		return nullptr;
	}

	// Load .NET Core
	void* load_assembly_and_get_function_pointer = nullptr;
	hostfxr_handle cxt = nullptr;
	int rc = init_fptr(configFile.c_str(), nullptr, &cxt);
	if (!STATUS_CODE_SUCCEEDED(rc) || cxt == nullptr)
	{
		if (rc == CoreHostIncompatibleConfig)
		{
			std::wstring msg = L"The required .NET 6 runtime is incompatible with the runtime " + get_loaded_runtime_version() + L" already loaded in the process.\n\nYou can try to disable other Excel add-ins to resolve the conflict.";
			ShowHostError(msg);
		}
		else
		{
			std::wstringstream stream;
			stream << L".NET runtime initialization failed: " << std::hex << std::showbase << rc << std::endl << std::endl;
			stream << L"You can find more information about this error at https://github.com/dotnet/runtime/blob/main/docs/design/features/host-error-codes.md";
			ShowHostError(stream.str());
		}
		close_fptr(cxt);
		return nullptr;
	}

	// Get the load assembly function pointer
	rc = get_delegate_fptr(
		cxt,
		hdt_load_assembly_and_get_function_pointer,
		&load_assembly_and_get_function_pointer);
	if (!STATUS_CODE_SUCCEEDED(rc) || load_assembly_and_get_function_pointer == nullptr)
	{
		std::wstringstream stream;
		stream << "Get .NET runtime delegate failed: " << std::hex << std::showbase << rc << std::endl << std::endl;
		stream << L"You can find more information about this error at https://github.com/dotnet/runtime/blob/main/docs/design/features/host-error-codes.md";
		ShowHostError(stream.str());
		close_fptr(cxt);
		return nullptr;
	}

	close_fptr(cxt);
	return (load_assembly_and_get_function_pointer_fn)load_assembly_and_get_function_pointer;
}
