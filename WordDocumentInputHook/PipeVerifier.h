#pragma once
#include <windows.h>
#include <string>
#include <thread>


class PipeVerifier
{
private:
    static const char* PIPE_NAME;
    DWORD currentPid = NULL;
public:
    PipeVerifier() {
        updatePid();
    }
    ~PipeVerifier() {

    }

    BOOL verify(bool exitCondition) {
        return connectToPipeInjectorAndVerify(exitCondition);
    }

    auto getPid() const {
        return currentPid;
    }
    void updatePid() {
        currentPid = GetCurrentProcessId();
    }
private:
    // 管道客户端连接函数
    BOOL connectToPipeInjectorAndVerify(bool exitCondition)
    {
        HANDLE hPipe = INVALID_HANDLE_VALUE;
        // 尝试连接管道服务器
        while (!exitCondition) {
            hPipe = CreateFileA(
                PIPE_NAME,
                GENERIC_READ | GENERIC_WRITE,
                0,              // 不共享
                NULL,           // 默认安全属性
                OPEN_EXISTING,  // 必须已存在
                0,              // 默认属性
                NULL            // 不指定模板文件
            );

            if (hPipe != INVALID_HANDLE_VALUE) {
                break;  // 连接成功
            }

            DWORD error = GetLastError();
            if (error != ERROR_PIPE_BUSY) {
                return FALSE;  // 连接失败
            }

            // 管道忙，等待后重试
            if (!WaitNamedPipeA(PIPE_NAME, 1000)) {  // 等待n毫秒
                return FALSE;
            }
        }
        if (exitCondition)
            return FALSE;

        // 设置管道读写模式
        DWORD mode = PIPE_READMODE_MESSAGE;
        if (!SetNamedPipeHandleState(hPipe, &mode, NULL, NULL)) {
            CloseHandle(hPipe);
            return FALSE;
        }
        BOOL success = FALSE;
        // 发送当前进程ID给注入器验证
        DWORD bytesWritten;
        if (WriteFile(hPipe, &currentPid, sizeof(DWORD), &bytesWritten, NULL)) {
            // 读取注入器的确认信号
            BOOL confirmation = FALSE;
            DWORD bytesRead;
            if (ReadFile(hPipe, &confirmation, sizeof(BOOL), &bytesRead, NULL)) {
                success = (confirmation == TRUE);
            }
        }
        CloseHandle(hPipe);
        return success;
    }

};

const char* PipeVerifier::PIPE_NAME = "\\\\.\\pipe\\FakeWordMfc_SlackingOff_Injector_Pipe";