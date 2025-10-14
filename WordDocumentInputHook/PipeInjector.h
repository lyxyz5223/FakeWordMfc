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

    // ���̲�ƥ�䷵��E_ACCESSDENIED���յ��ֽ�����һ�·���E_UNEXPECTED������ƥ�䷵��S_OK
    HRESULT injectAndWaitForConnection()
    {
        HRESULT hr = S_OK;
        // �����ܵ�������
        hr = startPipeServer();
        if (FAILED(hr))
            return hr;
        // ע��DLL��Ŀ�����
        hr = injectByPid(targetPid);
        if (FAILED(hr))
            return hr;
        // �ȴ�DLL���Ӳ���֤
        hr = waitForDLLConnection();
        return hr;
    }

private:
    // ���������Զ����õ�������
    void cleanup()
    {
        if (hPipe != INVALID_HANDLE_VALUE) {
            DisconnectNamedPipe(hPipe);
            CloseHandle(hPipe);
            hPipe = INVALID_HANDLE_VALUE;
        }
    }
    // �ڲ�ʹ�õ�ע�뺯��
    HRESULT injectByPid(DWORD pid)
    {
        std::wstring dllCmdLine;
        {// ȷ��������ڱ����ں��治�ᱻ����
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

        // ��Ŀ�����
        HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, pid);
        if (!hProcess)
            return HRESULT_FROM_WIN32(GetLastError());

        // ��Ŀ����̷����ڴ�
        size_t len = (dllCmdLine.size() + 1) * sizeof(wchar_t);
        LPVOID pRemote = VirtualAllocEx(hProcess, nullptr, len, MEM_COMMIT, PAGE_READWRITE);
        if (!pRemote) {
            CloseHandle(hProcess);
            return HRESULT_FROM_WIN32(GetLastError());
        }

        // д��dll·��
        if (!WriteProcessMemory(hProcess, pRemote, dllCmdLine.c_str(), len, nullptr)) {
            VirtualFreeEx(hProcess, pRemote, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return HRESULT_FROM_WIN32(GetLastError());
        }

        // ��ȡLoadLibraryW��ַ
        HMODULE mKernel32 = GetModuleHandleW(L"kernel32.dll");
        if (!mKernel32)
            return HRESULT_FROM_WIN32(GetLastError());
        LPTHREAD_START_ROUTINE pLoadLibraryW = (LPTHREAD_START_ROUTINE)GetProcAddress(mKernel32, "LoadLibraryW");
        if (!pLoadLibraryW)
            return HRESULT_FROM_WIN32(GetLastError());

        // ����Զ���̼߳���DLL
        HANDLE hThread = CreateRemoteThread(hProcess, nullptr, 0, pLoadLibraryW, pRemote, 0, nullptr);
        if (!hThread) {
            VirtualFreeEx(hProcess, pRemote, 0, MEM_RELEASE);
            CloseHandle(hProcess);
            return HRESULT_FROM_WIN32(GetLastError());
        }

        // �ȴ��߳̽���
        WaitForSingleObject(hThread, INFINITE);

        // ����
        VirtualFreeEx(hProcess, pRemote, 0, MEM_RELEASE);
        CloseHandle(hThread);
        CloseHandle(hProcess);

        return S_OK;
    }
    HRESULT startPipeServer()
    {
        // ���������ܵ�������
        hPipe = CreateNamedPipeA(
            PIPE_NAME,
            PIPE_ACCESS_DUPLEX,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            PIPE_UNLIMITED_INSTANCES,
            1024,  // �����������С
            1024,  // ���뻺������С
            10000,  // �ͻ��˳�ʱʱ�䣨10�룩
            NULL
        );

        if (hPipe == INVALID_HANDLE_VALUE) {
            auto code = GetLastError();
            std::cerr << "�����ܵ�ʧ��: " << code << std::endl;
            return HRESULT_FROM_WIN32(code);
        }

        std::cout << "�ܵ����������������ȴ�DLL����..." << std::endl;
        return S_OK;
    }

    HRESULT waitForDLLConnection()
    {
        // �ȴ�DLL�ͻ�������
        if (!ConnectNamedPipe(hPipe, NULL)) {
            DWORD error = GetLastError();
            if (error != ERROR_PIPE_CONNECTED) {
                std::cerr << "�ȴ�����ʧ��: " << error << std::endl;
                return HRESULT_FROM_WIN32(error);
            }
        }

        std::cout << "DLL�����ӣ���֤���..." << std::endl;

        // ����DLL���͵Ľ���ID
        DWORD receivedPid;
        DWORD bytesRead;

        if (!ReadFile(hPipe, &receivedPid, sizeof(DWORD), &bytesRead, NULL)) {
            auto error = GetLastError();
            std::cerr << "��ȡ����ʧ��: " << error << std::endl;
            return HRESULT_FROM_WIN32(error);
        }

        if (bytesRead != sizeof(DWORD)) {
            std::cerr << "�������ݴ�С����ȷ" << std::endl;
            return E_UNEXPECTED;
        }

        // ��֤����ID�Ƿ�ƥ��
        if (receivedPid != targetPid) {
            std::cerr << "����ID��ƥ��: ����=" << targetPid
                << ", ʵ��=" << receivedPid << std::endl;
            return E_ACCESSDENIED;
        }

        std::cout << "�����֤�ɹ���Ŀ��PID: " << receivedPid << std::endl;

        // ����ȷ���źŸ�DLL
        BOOL confirmation = TRUE;
        DWORD bytesWritten;
        if (!WriteFile(hPipe, &confirmation, sizeof(BOOL), &bytesWritten, NULL)) {
            auto error = GetLastError();
            std::cerr << "����ȷ��ʧ��: " << error << std::endl;
            return error;
        }

        return S_OK;
    }


};

const char* PipeInjector::PIPE_NAME = "\\\\.\\pipe\\FakeWordMfc_SlackingOff_Injector_Pipe";