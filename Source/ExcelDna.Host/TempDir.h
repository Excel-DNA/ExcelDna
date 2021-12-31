#pragma once
#include <string>

class TempDir
{
public:
	TempDir();
	~TempDir();
	const std::wstring& GetPath() const;

private:
	std::wstring  path;
};
