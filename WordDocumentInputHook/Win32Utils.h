#pragma once
#include <string>
#include <Windows.h>
#include <tlhelp32.h>

std::wstring GetDllPath(HMODULE hModule)
{
	std::wstring dllPath;
	int pFileNameLen = MAX_PATH;
	WCHAR* pFileName = new WCHAR[pFileNameLen];
	for (DWORD gmfRst = 0; (gmfRst = GetModuleFileName(hModule, pFileName, pFileNameLen)) != 0 && gmfRst == pFileNameLen;)
	{
		delete[] pFileName;
		pFileNameLen += MAX_PATH;
		pFileName = new WCHAR[pFileNameLen];
	}
	dllPath = pFileName;
	delete[] pFileName;
	return dllPath;
}

std::wstring HResultToWString(HRESULT hr)
{
	WCHAR* lpBuffer = new WCHAR[sizeof(WCHAR*) * 8];
	// ��ȡ������Ϣ�ĳ���
	DWORD messageSize = FormatMessage(
		FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		lpBuffer,
		0,
		NULL);
	delete[] lpBuffer;
	lpBuffer = new WCHAR[messageSize];
	//MessageBox(0, std::to_wstring(messageSize).c_str(), L"messageSize", 0);
	// ��������Ϣ���Ƶ��ַ���
	FormatMessage(
		FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		lpBuffer,
		messageSize,
		NULL);
	std::wstring message = lpBuffer;
	delete[] lpBuffer;
	return message;
}

DWORD GetMainThreadIdByProcessId(DWORD pid)
{
	DWORD threadId = 0;
	HANDLE hThreadSnap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
	if (hThreadSnap != INVALID_HANDLE_VALUE)
	{
		THREADENTRY32 tEntry32;
		tEntry32.dwSize = sizeof(THREADENTRY32);
		if (Thread32First(hThreadSnap, &tEntry32))
		{
			do {
				if (tEntry32.th32OwnerProcessID == pid)
				{
					threadId = tEntry32.th32ThreadID;
					break;
				}
			} while (Thread32Next(hThreadSnap, &tEntry32));
		}
		CloseHandle(hThreadSnap);
	}
	return threadId;
}
