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
    // �ܵ��ͻ������Ӻ���
    BOOL connectToPipeInjectorAndVerify(bool exitCondition)
    {
        HANDLE hPipe = INVALID_HANDLE_VALUE;
        // �������ӹܵ�������
        while (!exitCondition) {
            hPipe = CreateFileA(
                PIPE_NAME,
                GENERIC_READ | GENERIC_WRITE,
                0,              // ������
                NULL,           // Ĭ�ϰ�ȫ����
                OPEN_EXISTING,  // �����Ѵ���
                0,              // Ĭ������
                NULL            // ��ָ��ģ���ļ�
            );

            if (hPipe != INVALID_HANDLE_VALUE) {
                break;  // ���ӳɹ�
            }

            DWORD error = GetLastError();
            if (error != ERROR_PIPE_BUSY) {
                return FALSE;  // ����ʧ��
            }

            // �ܵ�æ���ȴ�������
            if (!WaitNamedPipeA(PIPE_NAME, 1000)) {  // �ȴ�n����
                return FALSE;
            }
        }
        if (exitCondition)
            return FALSE;

        // ���ùܵ���дģʽ
        DWORD mode = PIPE_READMODE_MESSAGE;
        if (!SetNamedPipeHandleState(hPipe, &mode, NULL, NULL)) {
            CloseHandle(hPipe);
            return FALSE;
        }
        BOOL success = FALSE;
        // ���͵�ǰ����ID��ע������֤
        DWORD bytesWritten;
        if (WriteFile(hPipe, &currentPid, sizeof(DWORD), &bytesWritten, NULL)) {
            // ��ȡע������ȷ���ź�
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