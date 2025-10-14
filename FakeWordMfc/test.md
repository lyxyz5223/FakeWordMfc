```cpp
#include <windows.h>
#include <ole2.h>
#include <olectl.h>
#include <msword9.h> // ������Word�汾��ͷ�ļ�

// �¼��������ӿ�ʵ��
class CWordEventSink : public IConnectionPointContainer
{
public:
    CWordEventSink() : m_cRef(1), m_pConnectionPoint(NULL), m_dwCookie(0) {}
    
    // IUnknown�ӿ�
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
    
    // IDispatch�ӿ�
    STDMETHODIMP GetTypeInfoCount(UINT* pctinfo) { return E_NOTIMPL; }
    STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) { return E_NOTIMPL; }
    STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, 
                             LCID lcid, DISPID* rgDispId) { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, 
                       WORD wFlags, DISPPARAMS* pDispParams, 
                       VARIANT* pVarResult, EXCEPINFO* pExcepInfo, 
                       UINT* puArgErr) {
        // ʵ�־�����¼�����
        switch (dispIdMember) {
            case 1: // DocumentChange�¼�
                // �����ĵ��仯
                break;
            // �����¼�����...
        }
        return S_OK;
    }
    
    // IConnectionPointContainer�ӿ�
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

// ���ӵ�Word�¼�
HRESULT ConnectToWordEvents()
{
    // ��ʼ��COM
    CoInitialize(NULL);
    
    // ����WordӦ�ó������
    IDispatch* pWordApp = NULL;
    HRESULT hr = CoCreateInstance(CLSID_WordApplication, NULL, CLSCTX_LOCAL_SERVER, 
                                 IID_IDispatch, (void**)&pWordApp);
    
    if (SUCCEEDED(hr)) {
        // ��ȡ���ӵ�����
        IConnectionPointContainer* pCPC = NULL;
        hr = pWordApp->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
        
        if (SUCCEEDED(hr)) {
            // �������ӵ�
            IConnectionPoint* pCP = NULL;
            hr = pCPC->FindConnectionPoint(IID__ApplicationEvents, &pCP);
            
            if (SUCCEEDED(hr)) {
                // �����¼�������
                CWordEventSink* pSink = new CWordEventSink();
                
                // ��������
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

// �Ͽ�����
void DisconnectFromWordEvents()
{
    if (m_dwCookie != 0) {
        // ��ȡ���ӵ�
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

�ؼ��㣺
1. ʵ��������COM�ӿ�
2. ��ȷ�������ü���
3. ʹ����ȷ�����ӵ�IID��_ApplicationEvents�ȣ�
4. �������ӵ�Cookie���ڶϿ�����