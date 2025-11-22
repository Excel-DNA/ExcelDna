#include "path.h"

void RemoveFileSpecFromPath(std::wstring& filePath)
{
	size_t dirSepInd = filePath.find_last_of(L'\\');
	if (dirSepInd >= 0)
		filePath.erase(dirSepInd);
}

void StripPath(std::wstring& filePath)
{
	size_t dirSepInd = filePath.find_last_of(L'\\');
	if (dirSepInd >= 0)
		filePath.erase(0, dirSepInd + 1);
}

std::wstring GetDirectory(const std::wstring& filePath)
{
	std::wstring result(filePath);
	RemoveFileSpecFromPath(result);
	return result;
}

std::wstring GetDirectoryName(const std::wstring& filePath)
{
	std::wstring result(GetDirectory(filePath));
	StripPath(result);
	return result;
}

std::wstring GetFileName(const std::wstring& filePath)
{
	std::wstring result(filePath);
	StripPath(result);
	return result;
}

void RemoveExtension(std::wstring& filePath)
{
	size_t dotInd = filePath.find_last_of(L'.');
	filePath.erase(dotInd);
}

void RenameExtension(std::wstring& filePath, std::wstring ext)
{
	size_t dotInd = filePath.find_last_of(L'.');
	filePath.replace(dotInd, std::wstring::npos, ext);
}


std::wstring PathCombine(const std::wstring& path1, const std::wstring& path2)
{
	return path1 + L"\\" + path2;
}

std::wstring PathCombine(const std::wstring& path1, const std::wstring& path2, const std::wstring& path3)
{
	return PathCombine(PathCombine(path1, path2), path3);
}
