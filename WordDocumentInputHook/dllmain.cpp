// dllmain.cpp : 定义 DLL 应用程序的入口点。
#include "pch.h"
#include "dllmain.h"
#include <string>
#include <vector>
#include <functional>

#include "PipeInjector.h"
#include "PipeVerifier.h"
#include "StringProcess.h"
#include "MessageText.h"
#include "ShareMemory.h"
#include "Win32Utils.h"
#include "FileLogger.h"
#include <shellapi.h>
#include <mutex>
#include <shared_mutex>
#include <imm.h>
#pragma comment(lib, "Imm32.lib")
#include "TSFManager.h"
// 消息队列获取的消息
#define WM_TSF_INIT WM_USER + 1037 // TSF初始化消息
// 进程中的全局单例
HMODULE gModule = NULL;
//#define SHAREMEM_INSTANCE_LOCK_MODULE_NAME L"Global\\FakeWordMfc_SlackingOff_Instance_Lock"
//HANDLE hModuleLockMapFile = 0;

// 注入后使用的一些变量
struct RemoteInfoStruct {
	HWND windowId = 0; // 当前窗口句柄
	std::wstring injectId;
	DWORD processId = 0; // 当前进程id
	DWORD windowThreadId = 0; // 窗口线程id
} gRemoteInfo;
#define SHAREMEM_GLOBAL_REMOTEINFO_PREFIX L"Global\\FakeWordMfc_SlackingOff_RemoteInfo"
#define SHAREMEM_GLOBAL_REMOTEINFO_MTX_PREFIX L"Global\\FakeWordMfc_SlackingOff_RemoteInfo_Mtx"
// 注入后使用的一些全局变量
CTSFManager* gTSFMgr = nullptr;
// 窗口子类化，暂存变量
WNDPROC gWndProcOld = nullptr; // 临时保存旧的窗口过程
LRESULT CALLBACK SubClassWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam); // 子类化后的窗口过程

//#pragma push_macro("LOCK")
//#undef LOCK
//#define LOCK(mtx) std::unique_lock<std::mutex> lock{ mtx }
//#undef LOCK
//#pragma pop_macro("LOCK")

class ThreadSafeWString {
private:
	std::wstring data;
	mutable std::shared_mutex rwMtx;  // 读写锁

public:
	using size_type = std::wstring::size_type;

	// 构造函数
	ThreadSafeWString() {

	}

	template<typename T>
	ThreadSafeWString(T&& t) : data(std::forward<T>(t)) {
	}

	// 写操作 - 需要独占锁（排他锁）
	template<typename T>
	ThreadSafeWString& operator=(T&& t) {
		std::unique_lock<std::shared_mutex> lock(rwMtx);  // 独占锁，阻塞所有读写
		data = std::forward<T>(t);
		return *this;
	}

	template<typename T>
	ThreadSafeWString& operator+=(T&& t) {
		std::unique_lock<std::shared_mutex> lock(rwMtx);  // 独占锁
		data += std::forward<T>(t);
		return *this;
	}
	
	ThreadSafeWString& operator=(ThreadSafeWString& t) {
		std::unique_lock<std::shared_mutex> lock(rwMtx);  // 独占锁，阻塞所有读写
		data = t.data;
		return *this;
	}

	ThreadSafeWString& operator+=(ThreadSafeWString& t) {
		std::unique_lock<std::shared_mutex> lock(rwMtx);  // 独占锁
		data += t.data;
		return *this;
	}

	void push_back(wchar_t c) {
		std::unique_lock<std::shared_mutex> lock(rwMtx);  // 独占锁
		data.push_back(c);
	}

	void clear() {
		std::unique_lock<std::shared_mutex> lock(rwMtx);  // 独占锁
		data.clear();
	}

	// 读操作 - 使用共享锁，允许多线程同时读取
	wchar_t operator[](size_type idx) const {
		std::shared_lock<std::shared_mutex> lock(rwMtx);  // 共享锁，允许多个读取者
		return data[idx];
	}

	wchar_t at(size_type idx) const {
		std::shared_lock<std::shared_mutex> lock(rwMtx);  // 共享锁
		if (idx >= data.size()) {
			throw std::out_of_range("ThreadSafeWString::at");
		}
		return data[idx];
	}

	size_type size() const {
		std::shared_lock lock(rwMtx);  // 共享锁
		return data.size();
	}

	bool empty() const {
		std::shared_lock lock(rwMtx);  // 共享锁
		return data.empty();
	}

	// 获取字符串副本（线程安全）
	std::wstring str() const {
		std::shared_lock lock(rwMtx);  // 共享锁
		return data;
	}

	// 查找操作（读操作）
	size_type find(const std::wstring& str, size_type pos = 0) const {
		std::shared_lock lock(rwMtx);  // 共享锁
		return data.find(str, pos);
	}

	// 子字符串（读操作）
	std::wstring substr(size_type pos = 0, size_type count = std::wstring::npos) const {
		std::shared_lock lock(rwMtx);  // 共享锁
		return data.substr(pos, count);
	}
};
//ThreadSafeWString gInjectId; // 全局注入id，共享内存
#define SHAREMEM_GLOBAL_INJECTID_NAME L"Global\\FakeWordMfc_SlackingOff_InjectId"
#define SHAREMEM_GLOBAL_INJECTID_MTX_NAME L"Global\\FakeWordMfc_SlackingOff_InjectId_Mtx"

// Word端
bool gbDaemonThreadExit = false;
std::thread gDaemonThread;

// Word端全局hook变量
HHOOK gHHookCallWnd = 0;
HHOOK gHHookGetMsg = 0; // 暂不使用
bool gIsCurrentWord = false;

// logger
FileLogger gFileLogger;


#define STOI_CATCH \
	catch (std::invalid_argument err) \
	{ \
		/* 参数错误 */ \
		gFileLogger << str2wstr_2UTF8(err.what()); \
	} \
	catch (std::out_of_range err) \
	{ \
		/* 超出数据范围 */ \
		gFileLogger << str2wstr_2UTF8(err.what()); \
	} \
	catch (...) \
	{ \
		/* 其他错误，unreachable code */ \
	} \
	static_assert(true, "")

// DLL入口函数
BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{

    switch (ul_reason_for_call)
    {
	case DLL_PROCESS_ATTACH: // 进程调用的DllMain，仅在进程第一次加载Dll时候调用
	{
		if (gModule)
			return FALSE; // 确保进程单例
		gModule = hModule;

		// 下面代码用于在操作系统周期内是单例的，目前缺少互斥锁（共享内存实现）防止竞争访问
		//hModuleLockMapFile = OpenFileMapping(
		//	FILE_MAP_ALL_ACCESS,                  // read/write access
		//	FALSE,                                // do not inherit the name
		//	SHAREMEM_INSTANCE_LOCK_MODULE_NAME);  // name of mapping object
		//if (!hModuleLockMapFile)
		//{
		//	hModuleLockMapFile = CreateFileMappingW(
		//		INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(HMODULE), SHAREMEM_INSTANCE_LOCK_MODULE_NAME);
		//	if (hModuleLockMapFile)
		//	{
		//		HMODULE* smhModule = (HMODULE*)MapViewOfFile(hModuleLockMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(HMODULE));
		//		if (smhModule)
		//		{
		//			CopyMemory(smhModule, hModule, sizeof(HMODULE));
		//			UnmapViewOfFile(smhModule);
		//		}
		//	}
		//	else
		//		return FALSE;
		//}
		//else
		//{
		//	// 如果成功打开，读取内容，如果内容为空，说明当前进程内没有运行dll
		//	HMODULE* smhModule = (HMODULE*)MapViewOfFile(hModuleLockMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(HMODULE));
		//	if (smhModule)
		//	{
		//		if (*smhModule)
		//		{
		//			// 当前dll已经被加载，则退出当前实例
		//			UnmapViewOfFile(smhModule);
		//			return FALSE;
		//		}
		//		else // 没有加载dll则保存当前实例
		//		{
		//			CopyMemory(smhModule, hModule, sizeof(HMODULE));
		//			UnmapViewOfFile(smhModule);
		//		}
		//	}
		//	else
		//		return FALSE;
		//}
		
		std::wstring injectId;
		if (!ReadShareMemory<ThreadSafeWString>(SHAREMEM_GLOBAL_INJECTID_NAME, [&injectId](ThreadSafeWString* iid) -> BOOL {
			injectId = iid->str();
			//MessageBox(0, (L"Current inject id: " + injectId).c_str(), L"", 0);
			return TRUE;
			}))
		{
			//MessageBox(0, L"无法映射共享内存: ThreadSafeWString", L"", MB_ICONERROR);
		}
		if (!ReadShareMemory<RemoteInfoStruct>(SHAREMEM_GLOBAL_REMOTEINFO_PREFIX + injectId, [](RemoteInfoStruct* ri) -> BOOL {
			gRemoteInfo.injectId = ri->injectId;
			gRemoteInfo.processId = ri->processId;
			gRemoteInfo.windowId = ri->windowId;
			gRemoteInfo.windowThreadId = ri->windowThreadId;
			//std::wstringstream ss;
			//ss << L"Current remote info: {\n\tinjectId: " << ri->injectId
			//	<< L"\n\tprocessId: " << ri->processId
			//	<< L"\n\twindowId: " << ri->windowId
			//	<< L"\n\twindowThreadId: " << ri->windowThreadId
			//	<< L"\n}";
			//MessageBox(0, ss.str().c_str(), L"", 0);
			return TRUE;
			}))
		{
			//MessageBox(0, L"无法映射共享内存: RemoteInfoStruct", L"", MB_ICONERROR);
		}
		//MessageBox(0, L"Process detach", L"", 0);

		//// 记录日志
		//if (!gFileLogger.open(GetDllPath(hModule) + L"." + gRemoteInfo.injectId + L".log", std::fstream::out))
		//{ // 日志打开失败
		//	MessageBox(0, L"无法打开日志文件", L"", MB_ICONERROR);
		//}

		// 已经通过共享内存传递
		//if (gRemoteInfo.windowId)
		//{
		//	DWORD tmpPid = 0;
		//	gRemoteInfo.windowThreadId = GetWindowThreadProcessId(gRemoteInfo.windowId, &tmpPid);
		//}


		// 判断当前dll在哪里加载，如果是Word，那么开始初始化，并进行Hook
		gDaemonThread = std::thread(daemonThreadProc, hModule); // 启动守护线程
	}
		break;
	case DLL_THREAD_ATTACH: // 任何线程初次加载Dll都会调用
		break;
	case DLL_THREAD_DETACH: // 线程卸载Dll
		break;
	case DLL_PROCESS_DETACH: // 进程卸载Dll
		//CloseHandle(hModuleLockMapFile);
		gbDaemonThreadExit = true; // 设置线程退出标志，让线程退出
		gModule = NULL; // 进程模块单例互斥锁清零
		ZeroMemory(&gRemoteInfo, sizeof(gRemoteInfo)); // dll数据内存设为0
		if (gDaemonThread.joinable())
			gDaemonThread.join(); // 等待守护进程退出
		if (gHHookCallWnd)
			UnhookWindowsHookEx(gHHookCallWnd); // 取消挂钩
		gIsCurrentWord = false; // 重置状态量：当前dll是否运行在word进程内
		if (gWndProcOld)
		{
			SetWindowLongPtr(gRemoteInfo.windowId, GWLP_WNDPROC, (LONG_PTR)gWndProcOld);
			gWndProcOld = nullptr;
		}
		if (gTSFMgr)
			delete gTSFMgr;
		gFileLogger.close(); // 关闭日志文件
        break;
    }
    return TRUE;
}

HRESULT InjectInitShareMemory(HWND hwnd, const char* injectId, DWORD threadId, DWORD processId, std::vector<HANDLE>& handles)
{
	HANDLE hMapFileInjectId = 0; // 由 CreateAndWriteShareMemory 写入，函数结束时候关闭句柄
	HANDLE hMapFileRemoteInfo = 0; // 由 CreateAndWriteShareMemory 写入，函数结束时候关闭句柄
	std::wstring ijn = SHAREMEM_GLOBAL_INJECTID_NAME;
	auto writeGlobalInjectIdShareMemory = [=](auto* s) -> BOOL {
		//gInjectId = *s;
		s = new(s) ThreadSafeWString(str2wstr_2UTF8(injectId));
		return TRUE;
		};
	//std::wstring ijnm = SHAREMEM_GLOBAL_INJECTID_MTX_NAME;
	if (!ReadShareMemory<ThreadSafeWString>(ijn, writeGlobalInjectIdShareMemory))
	{
		if (!CreateAndWriteShareMemory<ThreadSafeWString>(ijn, writeGlobalInjectIdShareMemory, hMapFileInjectId))
		{
			MessageBox(0, L"无法创建共享内存区域，请检查权限，或者尝试以管理员身份运行", L"", MB_ICONERROR);
			return E_ACCESSDENIED;
		}
	}
	std::wstring shareMemName = SHAREMEM_GLOBAL_REMOTEINFO_PREFIX;
	shareMemName += str2wstr_2UTF8(injectId);
	if (!CreateAndWriteShareMemory<RemoteInfoStruct>(shareMemName,
		[=](RemoteInfoStruct* ri) -> BOOL {
			ri->injectId = str2wstr_2UTF8(injectId);
			ri->windowId = hwnd;
			ri->processId = processId;
			ri->windowThreadId = threadId;
			return TRUE;
		}, hMapFileRemoteInfo)) {
		MessageBox(0, L"无法创建共享内存区域，请检查权限，或者尝试以管理员身份运行", L"", MB_ICONERROR);
		return E_ACCESSDENIED;
	}
	handles.push_back(hMapFileInjectId);
	handles.push_back(hMapFileRemoteInfo);
	return S_OK;
}

HRESULT InjectToProcess(DWORD pid, const char* injectId, HWND targetWindowId)
{
	if (!gModule)
		return ERROR_INVALID_STATE; // 全局模块句柄为空，异常！
	gbDaemonThreadExit = true; // 注入时先退出当前dll实例的校验线程，防止竞争导致校验失败
	if (gDaemonThread.joinable())
		gDaemonThread.join();
	PipeInjector injector(gModule, pid);
	return injector.injectAndWaitForConnection();
}

HRESULT InjectToProcess(DWORD pid, const char* injectId)
{
	std::vector<HANDLE> handles;
	HRESULT hr = InjectInitShareMemory(0, injectId, 0, pid, handles);
	if (FAILED(hr))
		return hr;
	hr = InjectToProcess(pid, injectId, 0);
	for (auto iter = handles.begin(); iter != handles.end(); iter++)
		CloseHandle(*iter);
	return hr;
}

HRESULT InjectToProcessByWindowHwnd(HWND hwnd, const char* injectId)
{
	std::vector<HANDLE> handles;
	DWORD processId = NULL;
	DWORD threadId = GetWindowThreadProcessId(hwnd, &processId);
	if (!threadId)
		return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
	HRESULT hr = InjectInitShareMemory(hwnd, injectId, threadId, processId, handles);
	if (FAILED(hr))
		return hr;
	hr = InjectToProcess(processId, injectId, hwnd);
	for (auto iter = handles.begin(); iter != handles.end(); iter++)
		CloseHandle(*iter);
	return hr;
}


void daemonThreadProc(HMODULE hModule) // 守护线程
{
	PipeVerifier pidVerifier;
	//MessageBox(0, (L"当前进程Id" + std::to_wstring(pidVerifier.getPid())).c_str(), L"", MB_ICONINFORMATION);
	if (pidVerifier.verify(gbDaemonThreadExit)) // 校验时候传入线程退出条件，一旦gbDaemonThreadExit为true，将停止校验，返回false
	{
		gIsCurrentWord = true;
		// 记录日志
		if (!gFileLogger.open(GetDllPath(hModule) + L"." + gRemoteInfo.injectId + L".log", std::fstream::out))
		{ // 日志打开失败
			MessageBox(0, L"无法打开日志文件", L"", MB_ICONERROR);
		}
		//std::wstringstream ss;
		//ss << L"Current remote info: {\n\tinjectId: " << gRemoteInfo.injectId
		//	<< L"\n\tprocessId: " << gRemoteInfo.processId
		//	<< L"\n\twindowId: " << gRemoteInfo.windowId
		//	<< L"\n\twindowThreadId: " << gRemoteInfo.windowThreadId
		//	<< L"\n}";
		//MessageBox(0, ss.str().c_str(), L"", 0);
		//MessageBox(0, (L"Inject OK! Current dll instance is running on process id: " + std::to_wstring(pidVerifier.getPid())).c_str(), L"OK", 0);
		gFileLogger << L"Inject OK! Current dll instance is running on process id: " + std::to_wstring(pidVerifier.getPid());
		// 校验完成后开始hook
		//DWORD threadMain = GetMainThreadIdByProcessId(pidVerifier.getPid());
		DWORD& windowThreadId = gRemoteInfo.windowThreadId; // 获取对于全局变量的引用
		if (!windowThreadId)
		{
			DWORD tmpPid = 0;
			// 如果窗口id为空（即未传参或为0），则使用备用方案获取
			windowThreadId = GetMainThreadIdByProcessId(gRemoteInfo.processId);
			if (!windowThreadId)
			{
				gFileLogger << L"获取窗口线程失败, 注入id: " << gRemoteInfo.injectId
					<< L", 窗口句柄: " << gRemoteInfo.windowId
					<< L", 进程id: " << gRemoteInfo.processId;
			}
		}
		//gHHook = SetWindowsHookEx(WH_GETMESSAGE, HookWorkGetMsgProc, hModule, windowThreadId);
		// 执行一次子类化
		if (!gWndProcOld)
		{
			gWndProcOld = (WNDPROC)SetWindowLongPtr(
				gRemoteInfo.windowId,
				GWLP_WNDPROC,
				(LONG_PTR)SubClassWndProc);
			if (gWndProcOld)
				gFileLogger << L"[daemonThreadProc] SetWindowLongPtr GWLP_WNDPROC SubClassWndProc successfully.";
			else
			{
				auto err = GetLastError();
				gFileLogger << L"[daemonThreadProc] SetWindowLongPtr GWLP_WNDPROC SubClassWndProc failed. HRESULT code: " << err << L", msg: " << HResultToWString(HRESULT_FROM_WIN32(err));
			}
		}
		gHHookCallWnd = SetWindowsHookEx(WH_CALLWNDPROC, HookWorkCallWndProc, hModule, windowThreadId);
		if (gHHookCallWnd)
		{
			//MessageBox(0, L"设置Hook成功", L"", MB_ICONINFORMATION);
			gFileLogger << L"设置Hook成功";
		}
		else
		{
			auto win32Error = GetLastError();
			auto errMsg = HResultToWString(win32Error);
			MessageBox(0, (L"设置Hook失败: " + errMsg).c_str(), L"", MB_ICONERROR);
			gFileLogger << L"设置Hook失败: " << errMsg;
		}
		PostMessage(gRemoteInfo.windowId, WM_TSF_INIT, 0, 0);
		while (!gbDaemonThreadExit)
		{
			//gHHook = SetWindowsHookEx(WH_GETMESSAGE, HookWorkGetMsgProc, hModule, windowThreadId);
			// 维持线程生命周期，直到dll卸载
			Sleep(100);
		}
	}
}





// 钩子过程 - 每当目标线程调用 GetMessage 或 PeekMessage 时，系统会调用此函数
LRESULT CALLBACK HookWorkGetMsgProc(int code, WPARAM wParam, LPARAM lParam)
{
	// 必须调用 CallNextHookEx
	if (code < 0) return CallNextHookEx(gHHookGetMsg, code, wParam, lParam);
	//if (code < 0 || code == HC_NOREMOVE)
	//	return CallNextHookEx(gHHook, code, wParam, lParam);
	// 处理消息
	MSG* msg = (MSG*)lParam;
	switch (msg->message)
	{
	case WM_IME_COMPOSITION:
	case WM_CHAR:
		gFileLogger << L"GetMessge: " << str2wstr_2UTF8(GetMessageName(msg->message));
		break;
	case WM_TSF_INIT:
	{
		// UI线程初始化tsf相关服务
		gFileLogger << L"(User)GetMessge: WM_TSF_INIT";
		//gTSFMgr->Initialize(gRemoteInfo.windowId);
		break;
	}
	default:
		break;
	}

	// 将消息传递给下一个钩子或目标窗口
	return CallNextHookEx(gHHookGetMsg, code, wParam, lParam);
}


// 使用调用窗口过程钩子
LRESULT CALLBACK HookWorkCallWndProc(int code, WPARAM wParam, LPARAM lParam)
{
	if (code < 0)
		return CallNextHookEx(gHHookCallWnd, code, wParam, lParam);

	CWPSTRUCT* cwp = (CWPSTRUCT*)lParam;
	switch (cwp->message)
	{
	case WM_IME_COMPOSITION:
		gFileLogger << L" WM_IME_COMPOSITION lParam: " << cwp->lParam;
		if (cwp->lParam & GCS_RESULTSTR)
		{
			auto hwnd = cwp->hwnd;
			HIMC hImc = ImmGetContext(hwnd); // 获取输入上下文
			if (hImc)
			{
				// 检查是否有结果字符串需要提交（用户确认了输入）
				if (lParam & GCS_RESULTSTR)
				{
					// 获取结果字符串的长度（字节数，包括null终止符）
					DWORD dwSize = ImmGetCompositionString(hImc, GCS_RESULTSTR, NULL, 0);
					if (dwSize > 0)
					{
						// 分配足够的内存来存储字符串
						// 注意：dwSize 是字节数，对于 Unicode 是字符数的2倍
						LPWSTR lpResultStr = new WCHAR[dwSize + 1];
						if (lpResultStr)
						{
							// 获取结果字符串
							DWORD dwRet = ImmGetCompositionString(hImc, GCS_RESULTSTR, lpResultStr, dwSize);
							if (dwRet > 0)
							{
								// 确保字符串正确终止
								lpResultStr[dwSize - 1] = L'\0';
								// 记录文本
								gFileLogger << L"GCS_RESULTSTR: [" << lpResultStr << L"]";
								// 通知IME我们已经处理了结果字符串，防止IME重复处理
								ImmNotifyIME(hImc, NI_COMPOSITIONSTR, CPS_COMPLETE, 0);
							}
							delete[] lpResultStr;
						}
					}
					else
					{
						if (dwSize == 0)
							gFileLogger << L"GCS_RESULTSTR: ";
						else if (dwSize == IMM_ERROR_NODATA)
							gFileLogger << L"GCS_RESULTSTR: IMM_ERROR_NODATA";
						else if (dwSize == IMM_ERROR_NODATA)
							gFileLogger << L"GCS_RESULTSTR: IMM_ERROR_NODATA";
						else
							gFileLogger << L"GCS_RESULTSTR: Unknown";
					}
				}

				// 同时处理进行中的组合
				//if (lParam & GCS_COMPSTR)
				//{
				//	// 获取组合字符串（未确认的输入，通常带下划线显示）
				//	DWORD dwSize = ImmGetCompositionString(hImc, GCS_COMPSTR, NULL, 0);
				//	if (dwSize > 0)
				//	{
				//		LPWSTR lpCompStr = (LPWSTR)malloc(dwSize + sizeof(WCHAR));
				//		if (lpCompStr)
				//		{
				//			ImmGetCompositionString(hImc, GCS_COMPSTR, lpCompStr, dwSize);
				//			lpCompStr[dwSize / sizeof(WCHAR)] = L'\0';
				//			gFileLogger << L"GCS_COMPSTR: [" << lpCompStr << L"]";
				//			// 显示组合字符串（通常带下划线）
				//			// DisplayCompositionString(lpCompStr);
				//			free(lpCompStr);
				//		}
				//	}
				//}

				ImmReleaseContext(hwnd, hImc); // 释放输入上下文
			}
		} // 不使用break;因为要fallthrough到日志打印
	case WM_IME_STARTCOMPOSITION:
	case WM_IME_ENDCOMPOSITION:
	case WM_CHAR:
	{
		gFileLogger << L"CallWndProc: " << str2wstr_2UTF8(GetMessageName(cwp->message));
		break;
	}
	case WM_TSF_INIT:
	{
		// UI线程初始化tsf相关服务
		gFileLogger << L"[CallWndProc] (User)WM_TSF_INIT";
		//if (!gTSFMgr)
		//{
		//	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
		//	if (FAILED(hr) && hr != RPC_E_CHANGED_MODE)
		//	{
		//		gFileLogger << L"[CallWndProc] (User)WM_TSF_INIT: CoInitializeEx failed: " << HResultToWString(hr);
		//		break;
		//	}
		//	HANDLE hTok;
		//	if (!OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, TRUE, &hTok))
		//	{// 一定进入此分支
		//		DWORD gle = GetLastError(); // 1008，当前线程没有身份令牌
		//		gFileLogger << L"OpenThreadToken fail GLE=" << gle;
		//	}
		//	else
		//	{// Unreachable code
		//		CloseHandle(hTok);
		//		gFileLogger << L"Thread token OK";
		//		gTSFMgr = new CTSFManager(gFileLogger);
		//		gTSFMgr->Initialize(gRemoteInfo.windowId);
		//	}
		//}
		break;
	}
	default:
		break;
	}


	return CallNextHookEx(gHHookCallWnd, code, wParam, lParam);
}

// 给当前线程临时绑定一个令牌（进程令牌复制而来）
HRESULT EnsureTokenOnce()
{
	HANDLE hProcTok = nullptr, hImp = nullptr;
	if (!OpenProcessToken(GetCurrentProcess(),
		TOKEN_DUPLICATE | TOKEN_IMPERSONATE | TOKEN_QUERY,
		&hProcTok))
		return HRESULT_FROM_WIN32(GetLastError());

	if (!DuplicateTokenEx(hProcTok, MAXIMUM_ALLOWED, nullptr,
		SecurityImpersonation, TokenImpersonation,
		&hImp))
	{
		CloseHandle(hProcTok);
		return HRESULT_FROM_WIN32(GetLastError());
	}

	// 绑到当前线程
	if (!SetThreadToken(nullptr, hImp))
	{
		CloseHandle(hImp);
		CloseHandle(hProcTok);
		return HRESULT_FROM_WIN32(GetLastError());
	}

	CloseHandle(hImp);
	CloseHandle(hProcTok);
	return S_OK;
}

void CreateTSFManager()
{
	HRESULT hr = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
	gFileLogger << L"[SubClassWndProc] (User)WM_TSF_INIT: CoInitializeEx HRESULT: " << hr << L", msg: " << HResultToWString(hr);
	if (FAILED(hr) && hr != RPC_E_CHANGED_MODE && hr != S_FALSE)
	{
		gFileLogger << L"[SubClassWndProc] (User)WM_TSF_INIT: CoInitializeEx failed: " << HResultToWString(hr);
	}
	else
	{
		//HRESULT hr = EnsureTokenOnce();
		//if (SUCCEEDED(hr))
		//{
		//	gFileLogger << L"[SubClassWndProc] EnsureTokenOnce OK";
		//	gTSFMgr = new CTSFManager(gFileLogger);
		//	gTSFMgr->Initialize(gRemoteInfo.windowId);
		//}
		//else
		//{
		//	_com_error err(hr);
		//	gFileLogger << L"[SubClassWndProc] EnsureTokenOnce failed: HRESULT code: " << hr << L", msg: " << err.ErrorMessage();
		//}
		gFileLogger << L"[SubClassWndProc] gRemoteInfo.windowId: " << gRemoteInfo.windowId;
		gTSFMgr = new CTSFManager(gFileLogger);
		SetForegroundWindow(gRemoteInfo.windowId);
		SetFocus(gRemoteInfo.windowId);
		//SetWindowLong(gRemoteInfo.windowId, GWL_STYLE, (WS_POPUP & (~WS_CHILD)) | WS_CAPTION | WS_TILEDWINDOW);
		gTSFMgr->Initialize(gRemoteInfo.windowId);
	}
}
LRESULT CALLBACK SubClassWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	auto rst = CallWindowProc(gWndProcOld, hWnd, msg, wParam, lParam);
	if (msg == WM_TSF_INIT)
	{
		if (!gTSFMgr)
			CreateTSFManager();
		else
			gTSFMgr->Initialize(hWnd);
		return 0; // 消息处理完毕
	}
	else {
		switch (msg)
		{
		case WM_SETFOCUS:
		{
			CreateTSFManager();
			if (gTSFMgr && hWnd == gRemoteInfo.windowId)
				gTSFMgr->Initialize(gRemoteInfo.windowId);
			break;
		}
		case WM_KILLFOCUS:
		{
			if (gTSFMgr && hWnd == gRemoteInfo.windowId)
				gTSFMgr->Uninitialize();
			break;
		}
		default:
			break;
		}
	}
	gFileLogger << L"[SubClassWndProc] hWnd: " << hWnd
		<< L", msg: " << GetMessageName(msg)
		<< L", wParam: " << wParam
		<< L", lParam: " << lParam;
	// 其他消息交给原窗口过程
	return rst;
}