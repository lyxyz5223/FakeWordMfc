#pragma once
#include <windows.h>
#include <iostream>
#include <string>

class PipeInjector {
private:
    static const char* PIPE_NAME;
    HANDLE hPipe = 0;
    DWORD targetPid = 0;
    HMODULE hModule = 0;

public:
    PipeInjector(HMODULE hModule, DWORD targetPid)
        : hPipe(INVALID_HANDLE_VALUE),
        targetPid(targetPid),
        hModule(hModule)
    {
    }
    ~PipeInjector() {
        cleanup();
    }

    // 进程不匹配返回E_ACCESSDENIED，收到字节数不一致返回E_UNEXPECTED，进程匹配返回S_OK
    HRESULT injectAndWaitForConnection()
    {
        HRESULT hr = S_OK;
        // 启动管道服务器
        hr = startPipeServer();
        if (FAILED(hr))
            return hr;
        // 注入DLL到目标进程
        hr = injectByPid(targetPid);
        if (FAILED(hr))
            return hr;
        // 等待DLL连接并验证
        hr = waitForDLLConnection();
        return hr;
    }

private:
    // 析构函数自动调用的清理函数
    void cleanup()
    {
        if (hPipe != INVALID_HANDLE_VALUE) {
            DisconnectNamedPipe(hPipe);
            CloseHandle(hPipe);
            hPipe = INVALID_HANDLE_VALUE;
        }
    }
    // 内部使用的注入函数
    HRESULT injectByPid(DWORD pid)
    {
        std::wstring dllCmdLine;
        {// 确保代码块内变量在后面不会被误用
            int pFileNameLen = MAX_PATH;
            WCHAR* pFileName = new WCHAR[pFileNameLen];
            for (DWORD gmfRst = 0; (gmfRst = GetModuleFileName(hModule, pFileName, pFileNameLen)) != 0 && gmfRst == pFileNameLen;)
            {
                delete[] pFileName;
                pFileNameLen += MAX_PATH;
                pFileName = new WCHAR[pFileNameLen];
            }
            dllCmdLine = pFileName;
            delete[] pFileName;
        }

        // 打开目标进程
        HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid);
        if (!hProcess)
            return HRESULT_FROM_WIN32(GetLastError());

        // 在目标进程分配内存
        size_t len = (dllCmdLine.size() + 1) * sizeof(wchar_t);
        LPVOID pRemote = VirtualAllocEx(hProcess, nullptr, len, MEM_COMMIT, PAGE_READWRITE);
        if (!pRemote) {
            CloseHandle(hProcess);
            return HRESULT_FROM_WIN32(GetLastError());
        }

        // 写入dll路径
        if (!WriteProcessMemory(hProcess, pRemote, dllCmdLine.c_str(), len, nullptr)) {
            VirtualFreeEx(hProcess, pRemote, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return HRESULT_FROM_WIN32(GetLastError());
        }

        // 获取LoadLibraryW地址
        HMODULE mKernel32 = GetModuleHandleW(L"kernel32.dll");
        if (!mKernel32)
            return HRESULT_FROM_WIN32(GetLastError());
        LPTHREAD_START_ROUTINE pLoadLibraryW = (LPTHREAD_START_ROUTINE)GetProcAddress(mKernel32, "LoadLibraryW");
        if (!pLoadLibraryW)
            return HRESULT_FROM_WIN32(GetLastError());

        // 创建远程线程加载DLL
        HANDLE hThread = CreateRemoteThread(hProcess, nullptr, 0, pLoadLibraryW, pRemote, 0, nullptr);
        if (!hThread) {
            VirtualFreeEx(hProcess, pRemote, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return HRESULT_FROM_WIN32(GetLastError());
        }

        // 等待线程结束
        WaitForSingleObject(hThread, INFINITE);

        // 清理
        VirtualFreeEx(hProcess, pRemote, 0, MEM_RELEASE);
        CloseHandle(hThread);
        CloseHandle(hProcess);

        return S_OK;
    }
    HRESULT startPipeServer()
    {
        // 创建命名管道服务器
        hPipe = CreateNamedPipeA(
            PIPE_NAME,
            PIPE_ACCESS_DUPLEX,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            PIPE_UNLIMITED_INSTANCES,
            1024,  // 输出缓冲区大小
            1024,  // 输入缓冲区大小
            10000,  // 客户端超时时间（10秒）
            NULL
        );

        if (hPipe == INVALID_HANDLE_VALUE) {
            auto code = GetLastError();
            std::cerr << "创建管道失败: " << code << std::endl;
            return HRESULT_FROM_WIN32(code);
        }

        std::cout << "管道服务器已启动，等待DLL连接..." << std::endl;
        return S_OK;
    }

    HRESULT waitForDLLConnection()
    {
        // 等待DLL客户端连接
        if (!ConnectNamedPipe(hPipe, NULL)) {
            DWORD error = GetLastError();
            if (error != ERROR_PIPE_CONNECTED) {
                std::cerr << "等待连接失败: " << error << std::endl;
                return HRESULT_FROM_WIN32(error);
            }
        }

        std::cout << "DLL已连接，验证身份..." << std::endl;

        // 接收DLL发送的进程ID
        DWORD receivedPid;
        DWORD bytesRead;

        if (!ReadFile(hPipe, &receivedPid, sizeof(DWORD), &bytesRead, NULL)) {
            auto error = GetLastError();
            std::cerr << "读取数据失败: " << error << std::endl;
            return HRESULT_FROM_WIN32(error);
        }

        if (bytesRead != sizeof(DWORD)) {
            std::cerr << "接收数据大小不正确" << std::endl;
            return E_UNEXPECTED;
        }

        // 验证进程ID是否匹配
        if (receivedPid != targetPid) {
            std::cerr << "进程ID不匹配: 期望=" << targetPid
                << ", 实际=" << receivedPid << std::endl;
            return E_ACCESSDENIED;
        }

        std::cout << "身份验证成功！目标PID: " << receivedPid << std::endl;

        // 发送确认信号给DLL
        BOOL confirmation = TRUE;
        DWORD bytesWritten;
        if (!WriteFile(hPipe, &confirmation, sizeof(BOOL), &bytesWritten, NULL)) {
            auto error = GetLastError();
            std::cerr << "发送确认失败: " << error << std::endl;
            return error;
        }

        return S_OK;
    }


};

const char* PipeInjector::PIPE_NAME = "\\\\.\\pipe\\FakeWordMfc_SlackingOff_Injector_Pipe";