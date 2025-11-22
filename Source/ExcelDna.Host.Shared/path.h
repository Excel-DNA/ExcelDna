#pragma once

#include <string>

void RemoveFileSpecFromPath(std::wstring& filePath);
void StripPath(std::wstring& filePath);
void RemoveExtension(std::wstring& filePath);
void RenameExtension(std::wstring& filePath, std::wstring ext);

std::wstring GetDirectory(const std::wstring& filePath);
std::wstring GetDirectoryName(const std::wstring& filePath);
std::wstring GetFileName(const std::wstring& filePath);

std::wstring PathCombine(const std::wstring& path1, const std::wstring& path2);
std::wstring PathCombine(const std::wstring& path1, const std::wstring& path2, const std::wstring& path3);
