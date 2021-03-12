#ifndef CSimpleScriptSite_H
#define CSimpleScriptSite_H

#include <atlbase.h>
#include <activscp.h>
#include <atlwin.h>

class CSimpleScriptSite :
    public IActiveScriptSite,
    public IActiveScriptSiteWindow
{
public:
    CSimpleScriptSite();
    virtual ~CSimpleScriptSite();

    // IUnknown

    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject);

    // IActiveScriptSite

    STDMETHOD(GetLCID)(LCID *plcid) { *plcid = 0; return S_OK; }
    STDMETHOD(GetItemInfo)(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti);
    STDMETHOD(GetDocVersionString)(BSTR *pbstrVersion) { *pbstrVersion = SysAllocString(L"1.0"); return S_OK; }
    STDMETHOD(OnScriptTerminate)(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo) { return S_OK; }
    STDMETHOD(OnStateChange)(SCRIPTSTATE ssScriptState) { return S_OK; }
    STDMETHOD(OnScriptError)(IActiveScriptError *pIActiveScriptError) { return S_OK; }
    STDMETHOD(OnEnterScript)(void) { return S_OK; }
    STDMETHOD(OnLeaveScript)(void) { return S_OK; }


    STDMETHOD(Eval)(const WCHAR* source, VARIANT* result);

    STDMETHOD(Run)(WCHAR* procname, DISPPARAMS* args, VARIANT* results);


    STDMETHOD(AddScript)(const WCHAR* source);

    // IActiveScriptSiteWindow

    STDMETHOD(GetWindow)(HWND *phWnd) { *phWnd = m_hWnd; return S_OK; }
    STDMETHOD(EnableModeless)(BOOL fEnable) { return S_OK; }

    // Miscellaneous

    HRESULT SetWindow(HWND hWnd) { m_hWnd = hWnd; return S_OK; }

    

public:
    LONG m_cRefCount;
    HWND m_hWnd;

	CComPtr<IActiveScript> spJScript;
	CComPtr<IActiveScriptParse> spJScriptParse;
};

#endif
