// This code was contributed by Koustubh Moharir (Thank you!)

#include "stdafx.h"
#include "MiscUtils.h"

TempFileHolder::~TempFileHolder()
{
	destroy();
}
void TempFileHolder::setFileName(std::wstring fileName)
{
	this->fileName = fileName;
}
void TempFileHolder::destroy()
{
	if (!this->fileName.empty())
	{
		::DeleteFile(fileName.c_str());
		this->fileName = L"";
	}
}

SafeVariantArray::SafeVariantArray(int size)
{
	pArray = SafeArrayCreateVector(VT_VARIANT, 0, size);
}

SafeVariantArray::~SafeVariantArray()
{
	SafeArrayDestroy(pArray);
}

void SafeVariantArray::lock()
{
	SafeArrayLock(pArray);
}
void SafeVariantArray::unlock()
{
	SafeArrayUnlock(pArray);
}
void SafeVariantArray::setElement(int i, INT32 value)
{
	auto& element = static_cast<VARIANT*>(pArray->pvData)[i];
	VariantInit(&element);
	element.vt = VT_I4;
	element.lVal = value;
}
void SafeVariantArray::setElement(int i, INT64 value)
{
	auto& element = static_cast<VARIANT*>(pArray->pvData)[i];
	VariantInit(&element);
	element.vt = VT_I8;
	element.llVal = value;
}
void SafeVariantArray::setElement(int i, BSTR value)
{
	auto& element = static_cast<VARIANT*>(pArray->pvData)[i];
	VariantInit(&element);
	element.vt = VT_BSTR;
	element.bstrVal = value;
}

SafeByteArray::SafeByteArray(void* data, int sizeInBytes)
{
	pArray = SafeArrayCreateVector(VT_UI1, 0, sizeInBytes);
	SafeArrayLock(pArray);
	memcpy(pArray->pvData, data, sizeInBytes);
	SafeArrayUnlock(pArray);
}
SafeByteArray::~SafeByteArray()
{
	SafeArrayDestroy(pArray);
}

bool FileExists(LPCTSTR szPath)
{
	DWORD dwAttrib = GetFileAttributes(szPath);

	return (dwAttrib != INVALID_FILE_ATTRIBUTES &&
		!(dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}

std::wstring LoadStringFromResource(HMODULE hModule, int id)
{
	const wchar_t* buffer = nullptr;
	LoadStringW(hModule, id, (LPWSTR)&buffer, 0);//ugly cast for badly designed API (specifying buffer size == 0 returns a read only pointer.
	return std::wstring(buffer, *((WORD*) buffer - 1));//The WORD preceding the address is the size of the resource string. which is not null-terminated.
}

std::wstring FormatString(std::wstring formatString, ...)
{
	LPWSTR buffer = NULL;
	va_list args = NULL;
	va_start(args, formatString);
	FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER,
		formatString.c_str(),
		0,
		0,
		(LPWSTR)&buffer,//ugly cast for badly designed API (FORMAT_MESSAGE_ALLOCATE_BUFFER returns a system allocated buffer)
		0,
		&args);

	va_end(args);

	std::wstring result(buffer ? buffer : L"");
	if (buffer)
		LocalFree(buffer);
	return result;
}

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

HRESULT HResultFromLastError()
{
	return HRESULT_FROM_WIN32(GetLastError());
}

SafeFile::SafeFile()
	: handle(INVALID_HANDLE_VALUE)
{

}

SafeFile::~SafeFile()
{
	if (handle != INVALID_HANDLE_VALUE)
		CloseHandle(handle);
}

HRESULT SafeFile::Create(std::wstring path, DWORD dwDesiredAccess, DWORD dwShareMode, DWORD dwCreationDisposition)
{
	handle = ::CreateFileW(path.c_str(), dwDesiredAccess, dwShareMode,
		NULL, dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, NULL);

	if (handle == INVALID_HANDLE_VALUE)
		return HResultFromLastError();
	return S_OK;
}

HRESULT SafeFile::Read(LPVOID pBuffer, DWORD nBufSize, DWORD& nBytesRead)
{
	BOOL b = ::ReadFile(handle, pBuffer, nBufSize, &nBytesRead, NULL);
	if (!b)
		return HResultFromLastError();

	return S_OK;
}

std::wstring UTF8toUTF16(const std::string& utf8)
{
	std::wstring utf16;
	int len = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, NULL, 0);
	if (len > 1)
	{
		utf16.resize(len);
		len = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, &utf16[0], len);
		utf16.resize(len);
	}
	return utf16;
}