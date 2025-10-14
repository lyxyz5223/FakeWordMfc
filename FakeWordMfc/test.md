```cpp
#include <windows.h>
#include <ole2.h>
#include <olectl.h>
#include <msword9.h> // 或其他Word版本的头文件

// 事件接收器接口实现
class CWordEventSink : public IConnectionPointContainer
{
public:
    CWordEventSink() : m_cRef(1), m_pConnectionPoint(NULL), m_dwCookie(0) {}
    
    // IUnknown接口
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) {
        if (riid == IID_IUnknown || riid == IID_IDispatch || riid == IID_IConnectionPointContainer) {
            *ppv = static_cast<IConnectionPointContainer*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    
    STDMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_cRef); }
    STDMETHODIMP_(ULONG) Release() {
        ULONG ref = InterlockedDecrement(&m_cRef);
        if (ref == 0) delete this;
        return ref;
    }
    
    // IDispatch接口
    STDMETHODIMP GetTypeInfoCount(UINT* pctinfo) { return E_NOTIMPL; }
    STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) { return E_NOTIMPL; }
    STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, 
                             LCID lcid, DISPID* rgDispId) { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, 
                       WORD wFlags, DISPPARAMS* pDispParams, 
                       VARIANT* pVarResult, EXCEPINFO* pExcepInfo, 
                       UINT* puArgErr) {
        // 实现具体的事件处理
        switch (dispIdMember) {
            case 1: // DocumentChange事件
                // 处理文档变化
                break;
            // 其他事件处理...
        }
        return S_OK;
    }
    
    // IConnectionPointContainer接口
    STDMETHODIMP EnumConnectionPoints(IEnumConnectionPoints** ppEnum) { return E_NOTIMPL; }
    STDMETHODIMP FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP) {
        if (riid == IID__ApplicationEvents) {
            if (!m_pConnectionPoint) {
                m_pConnectionPoint = new CConnectionPoint();
            }
            *ppCP = m_pConnectionPoint;
            return S_OK;
        }
        return CONNECT_E_NOCONNECTION;
    }
    
private:
    ULONG m_cRef;
    IConnectionPoint* m_pConnectionPoint;
    DWORD m_dwCookie;
};

// 连接到Word事件
HRESULT ConnectToWordEvents()
{
    // 初始化COM
    CoInitialize(NULL);
    
    // 创建Word应用程序对象
    IDispatch* pWordApp = NULL;
    HRESULT hr = CoCreateInstance(CLSID_WordApplication, NULL, CLSCTX_LOCAL_SERVER, 
                                 IID_IDispatch, (void**)&pWordApp);
    
    if (SUCCEEDED(hr)) {
        // 获取连接点容器
        IConnectionPointContainer* pCPC = NULL;
        hr = pWordApp->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
        
        if (SUCCEEDED(hr)) {
            // 查找连接点
            IConnectionPoint* pCP = NULL;
            hr = pCPC->FindConnectionPoint(IID__ApplicationEvents, &pCP);
            
            if (SUCCEEDED(hr)) {
                // 创建事件接收器
                CWordEventSink* pSink = new CWordEventSink();
                
                // 建立连接
                hr = pCP->Advise(pSink, &m_dwCookie);
                
                if (FAILED(hr)) {
                    pCP->Release();
                    pSink->Release();
                }
            }
            pCPC->Release();
        }
        pWordApp->Release();
    }
    
    return hr;
}

// 断开连接
void DisconnectFromWordEvents()
{
    if (m_dwCookie != 0) {
        // 获取连接点
        IConnectionPointContainer* pCPC = NULL;
        if (SUCCEEDED(pWordApp->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC))) {
            IConnectionPoint* pCP = NULL;
            if (SUCCEEDED(pCPC->FindConnectionPoint(IID__ApplicationEvents, &pCP))) {
                pCP->Unadvise(m_dwCookie);
                pCP->Release();
            }
            pCPC->Release();
        }
        
        m_dwCookie = 0;
    }
    
    CoUninitialize();
}
```

关键点：
1. 实现完整的COM接口
2. 正确处理引用计数
3. 使用正确的连接点IID（_ApplicationEvents等）
4. 保存连接点Cookie用于断开连接