#include "TempDir.h"
#include "utils.h"

#include <rpc.h>
#pragma comment(lib, "rpcrt4.lib")

namespace
{
	std::wstring CreateUuid()
	{
		UUID uuid;
		UuidCreate(&uuid);
		RPC_WSTR str;
		UuidToString(&uuid, &str);
		std::wstring result((TCHAR*)str);
		RpcStringFree(&str);
		return result;
	}
}

TempDir::TempDir()
{
	TCHAR lpTempPathBuffer[MAX_PATH];
	DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
	if (dwRetVal > MAX_PATH || (dwRetVal == 0))
		throw std::exception();

	path = lpTempPathBuffer + (L"\\DNA" + CreateUuid());
	CreateDirectory(path.c_str(), NULL);
}

TempDir::~TempDir()
{
	WIN32_FIND_DATA ffd;
	HANDLE hFind = FindFirstFile((path + L"\\*").c_str(), &ffd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
				DeleteFile((path + L"\\" + ffd.cFileName).c_str());
		} while (FindNextFile(hFind, &ffd) != 0);
		FindClose(hFind);
	}
	RemoveDirectory(path.c_str());
}

std::wstring TempDir::WriteFileBuf(const std::wstring& name, void* buf, DWORD size)
{
	std::wstring filePath = path + L"\\" + name;

	HANDLE hFile = CreateFile(filePath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		throw std::exception();

	DWORD dwBytesWritten;
	BOOL fSuccess = WriteFile(hFile, buf, size, &dwBytesWritten, NULL);
	if (!fSuccess)
	{
		CloseHandle(hFile);
		throw std::exception();
	}

	CloseHandle(hFile);
	return filePath;
}

