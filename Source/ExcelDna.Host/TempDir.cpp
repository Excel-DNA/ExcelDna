#include "TempDir.h"
#include "utils.h"

#include <cassert>
#include <rpc.h>
#pragma comment(lib, "rpcrt4.lib")

namespace
{
	std::wstring CreateUuid()
	{
		UUID uuid;
		RPC_STATUS s = UuidCreate(&uuid);
		assert((s == RPC_S_OK || s == RPC_S_UUID_LOCAL_ONLY) && "Failure: UuidCreate");

		RPC_WSTR str;
		s = UuidToString(&uuid, &str);
		assert(s == RPC_S_OK && "Failure: UuidToString");

		std::wstring result((TCHAR*)str);
		RpcStringFree(&str);
		return result;
	}
}

TempDir::TempDir()
{
	TCHAR lpTempPathBuffer[MAX_PATH];
	DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
	assert(!(dwRetVal > MAX_PATH || dwRetVal == 0) && "Failure: GetTempPath");

	path = PathCombine(lpTempPathBuffer, L"DNA" + CreateUuid());
	BOOL b = CreateDirectory(path.c_str(), NULL);
	assert(b && "Failure: CreateDirectory");
}

TempDir::~TempDir()
{
	WIN32_FIND_DATA ffd;
	HANDLE hFind = FindFirstFile(PathCombine(path, L"*").c_str(), &ffd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
				DeleteFile(PathCombine(path, ffd.cFileName).c_str());
		} while (FindNextFile(hFind, &ffd) != 0);
		FindClose(hFind);
	}
	RemoveDirectory(path.c_str());
}

const std::wstring& TempDir::GetPath() const
{
	return path;
}
