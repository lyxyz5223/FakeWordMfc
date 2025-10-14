#pragma once
#include <string>
#include <functional>
#include <Windows.h>
#include "Win32Utils.h"

template <typename T>
BOOL ReadShareMemory(std::wstring name, std::function<BOOL(T*)> successExecFunc)
{
	HANDLE remoteInfoShareMemoryMapFile = OpenFileMapping(
		FILE_MAP_ALL_ACCESS,                  // read/write access
		FALSE,                                // do not inherit the name
		name.c_str());  // name of mapping object
	if (!remoteInfoShareMemoryMapFile)
	{
		return FALSE; // 读取信息失败
	}
	else
	{
		// 如果成功打开，读取信息，保存到新位置
		auto hMapView = (T*)MapViewOfFile(remoteInfoShareMemoryMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(T));
		if (hMapView)
		{
			BOOL rst = successExecFunc(hMapView); // 成功回调
			UnmapViewOfFile(hMapView);
			CloseHandle(remoteInfoShareMemoryMapFile);
			return rst;
		}
		CloseHandle(remoteInfoShareMemoryMapFile);
		return FALSE;
	}
	return TRUE; // unreachable code
}

template <typename T>
BOOL CreateAndWriteShareMemory(std::wstring name, std::function<BOOL(T*)> writeFunc, HANDLE& shareMemoryHandle)
{
	// 先检测是否已经创建了共享内存区域，如果没有创建则先创建，否则直接使用
	HANDLE hMapFile = OpenFileMapping(
		FILE_MAP_ALL_ACCESS,                  // read/write access
		FALSE,                                // do not inherit the name
		name.c_str());  // name of mapping object
	if (!hMapFile)
	{
		hMapFile = CreateFileMappingW(
			INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(T), name.c_str());
		if (!hMapFile)
		{
			auto win32Error = GetLastError();
			std::wstring errMsg = HResultToWString(HRESULT_FROM_WIN32(win32Error));
			return FALSE;
		}
	}
	T* hMapView = (T*)MapViewOfFile(hMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(T));
	if (!hMapView)
	{
		CloseHandle(hMapFile);
		return FALSE;
	}
	BOOL rst = writeFunc(hMapView);
	UnmapViewOfFile(hMapView);
	shareMemoryHandle = hMapFile; // 交给外部管理共享内存的句柄关闭
	return rst;
}
