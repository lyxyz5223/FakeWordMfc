#pragma once
#include <Windows.h>

extern "C" {
	_declspec(dllexport) HRESULT InjectToProcess(DWORD pid, const char* injectId);
	_declspec(dllexport) HRESULT InjectToProcessByWindowHwnd(HWND hwnd, const char* injectId);
}

LRESULT CALLBACK HookWorkGetMsgProc(int code, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK HookWorkCallWndProc(int code, WPARAM wParam, LPARAM lParam);
void daemonThreadProc(HMODULE hModule);