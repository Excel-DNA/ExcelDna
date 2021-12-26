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

#include <cassert>
#include <iostream>
#include <string>

#include "TempDir.h"

using string_t = std::basic_string<char_t>;

// Globals to hold hostfxr exports
hostfxr_initialize_for_runtime_config_fn init_fptr;
hostfxr_get_runtime_delegate_fn get_delegate_fptr;
hostfxr_close_fn close_fptr;

// Forward declarations
bool load_hostfxr();
load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly(const char_t* assembly);

// Provide a callback for any catastrophic failures.
// The provided callback will be the last call prior to a rude-abort of the process.
// void __stdcall set_failure_callback(failure_fn cb) {}

// Preload the runtime.
// The runtime is lazily loaded whenever the first export is called. This function
// preloads the runtime independent of calling any export and avoids the startup
// cost associated with calling an export for the first time.
void __stdcall preload_runtime(void) {}

TempDir tempDir;

// TODO: Might return the fn*
int load_runtime_and_run(LPCWSTR basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll)
{
	// Get the current executable's directory
	// This sample assumes the managed assembly to load and its runtime configuration file are next to the host
	char_t host_path[MAX_PATH];

	auto size = ::GetFullPathNameW(basePath, sizeof(host_path) / sizeof(char_t), host_path, nullptr);
	assert(size != 0);

	string_t root_path = host_path;
	auto pos = root_path.find_last_of('\\');
	assert(pos != string_t::npos);
	root_path = root_path.substr(0, pos + 1);

	//
	// STEP 1: Load HostFxr and get exported hosting functions
	//
	if (!load_hostfxr())
	{
		assert(false && "Failure: load_hostfxr()");
		return EXIT_FAILURE;
	}

	//
	// STEP 2: Initialize and start the .NET Core runtime
	//
	const string_t config_path = root_path + L"ExcelDna.Host.runtimeconfig.json";
	const string_t deps_path = root_path + L"ExcelDna.Host.deps.json";
	load_assembly_and_get_function_pointer_fn load_assembly_and_get_function_pointer = nullptr;
	const char_t* config_path_str = config_path.c_str();
	const char_t* deps_path_str = deps_path.c_str();
	load_assembly_and_get_function_pointer = get_dotnet_load_assembly(config_path_str);
	assert(load_assembly_and_get_function_pointer != nullptr && "Failure: get_dotnet_load_assembly()");

	//
	// STEP 3: TODO: Copy managed assembly from resources to some temp file
	//

	HRSRC hResManagedHost = FindResource(hModuleXll, L"EXCELDNA.MANAGEDHOST", L"ASSEMBLY");
	HGLOBAL hManagedHost = LoadResource(hModuleXll, hResManagedHost);
	void* buf = LockResource(hManagedHost);
	DWORD resSize = SizeofResource(hModuleXll, hResManagedHost);
	std::wstring hostFile = tempDir.WriteFileBuf(L"ExcelDna.ManagedHost.dll", buf, resSize);

	//
	// STEP 4: Load managed assembly and get function pointer to a managed method
	//
	const string_t dotnetlib_path = hostFile;
	const char_t* dotnet_type = L"ExcelDna.ManagedHost.AddInInitialize, ExcelDna.ManagedHost";
	const char_t* dotnet_type_method = L"Initialize";

	// Function pointer to managed delegate with non-default signature
	typedef short (CORECLR_DELEGATE_CALLTYPE* xladdin_initialize_fn)(void* xlAddInExportInfo, void* hModuleXLL, void* pPathXLL);
	xladdin_initialize_fn init = nullptr;
	int rc = load_assembly_and_get_function_pointer(
		dotnetlib_path.c_str(),
		dotnet_type,
		dotnet_type_method,
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		(void**)&init);
	assert(rc == 0 && init != nullptr && "Failure: load_assembly_and_get_function_pointer()");

	short res = init(pExportInfo, hModuleXll, (void*)pathXll);

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
bool load_hostfxr()
{
	// Pre-allocate a large buffer for the path to hostfxr
	char_t buffer[MAX_PATH];
	size_t buffer_size = sizeof(buffer) / sizeof(char_t);
	int rc = get_hostfxr_path(buffer, &buffer_size, nullptr);
	if (rc != 0)
		return false;

	// Load hostfxr and get desired exports
	void* lib = load_library(buffer);
	init_fptr = (hostfxr_initialize_for_runtime_config_fn)get_export(lib, "hostfxr_initialize_for_runtime_config");
	get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)get_export(lib, "hostfxr_get_runtime_delegate");
	close_fptr = (hostfxr_close_fn)get_export(lib, "hostfxr_close");

	return (init_fptr && get_delegate_fptr && close_fptr);
}

// Load and initialize .NET Core and get desired function pointer for scenario
load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly(const char_t* config_path)
{
	std::string configText = R"({
  "runtimeOptions": {
    "tfm": "net5.0",
    "framework": {
      "name": "Microsoft.WindowsDesktop.App",
      "version": "5.0.1"
    }
  }
})";
	std::wstring configFile = tempDir.WriteFileBuf(L"ExcelDna.Host.runtimeconfig.json", (void*)configText.c_str(), configText.length());

	// Load .NET Core
	void* load_assembly_and_get_function_pointer = nullptr;
	hostfxr_handle cxt = nullptr;
	int rc = init_fptr(configFile.c_str(), nullptr, &cxt);
	if (rc != 0 || cxt == nullptr)
	{
		std::cerr << "Init failed: " << std::hex << std::showbase << rc << std::endl;
		close_fptr(cxt);
		return nullptr;
	}

	// Get the load assembly function pointer
	rc = get_delegate_fptr(
		cxt,
		hdt_load_assembly_and_get_function_pointer,
		&load_assembly_and_get_function_pointer);
	if (rc != 0 || load_assembly_and_get_function_pointer == nullptr)
		std::cerr << "Get delegate failed: " << std::hex << std::showbase << rc << std::endl;

	close_fptr(cxt);
	return (load_assembly_and_get_function_pointer_fn)load_assembly_and_get_function_pointer;
}
