#pragma once
#include <Windows.h>
#include <string>

class TempDir
{
public:
	TempDir();
	~TempDir();
	std::wstring WriteFileBuf(const std::wstring& name, void* buf, DWORD size);

private:
	std::wstring  path;
};
