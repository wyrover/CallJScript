
#include "SimpleScriptSite.h"

STDMETHODIMP_(ULONG) CSimpleScriptSite::AddRef()
{
    return InterlockedIncrement(&m_cRefCount);
}

STDMETHODIMP_(ULONG) CSimpleScriptSite::Release()
{
    if (!InterlockedDecrement(&m_cRefCount))
    {
        delete this;
        return 0;
    }
    return m_cRefCount;
}


CSimpleScriptSite::CSimpleScriptSite() : m_cRefCount(1), m_hWnd(NULL)
{
	
	spJScript.CoCreateInstance(OLESTR("JScript"));
	spJScript->SetScriptSite(this);
	spJScript->QueryInterface(&spJScriptParse);
	spJScriptParse->InitNew();
}


CSimpleScriptSite::~CSimpleScriptSite()
{
	spJScriptParse = NULL;
	spJScript = NULL;
	
}

STDMETHODIMP CSimpleScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
    if (riid == IID_IUnknown || riid == IID_IActiveScriptSiteWindow)
    {
        *ppvObject = (IActiveScriptSiteWindow *)this;
        AddRef();
        return NOERROR;
    }
    if (riid == IID_IActiveScriptSite)
    {
        *ppvObject = (IActiveScriptSite *)this;
        AddRef();
        return NOERROR;
    }
    return E_NOINTERFACE;
}


STDMETHODIMP CSimpleScriptSite::GetItemInfo(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown** ppiunkItem, ITypeInfo** ppti)
{
	return TYPE_E_ELEMENTNOTFOUND;
}

STDMETHODIMP CSimpleScriptSite::Eval(const WCHAR* source, VARIANT* result)
{
	if (!source)
		return E_POINTER;

	EXCEPINFO ei = { };

	return spJScriptParse->ParseScriptText(source, nullptr, nullptr, nullptr, 0, 0, SCRIPTTEXT_ISEXPRESSION, result, &ei);
}

STDMETHODIMP CSimpleScriptSite::Run(WCHAR* procname, DISPPARAMS* args, VARIANT* results)
{
	if (procname == nullptr)
		return E_POINTER;

	IDispatch* disp = nullptr;
	spJScript->GetScriptDispatch(nullptr, &disp);
	DISPID dispid = 0;
	disp->GetIDsOfNames(IID_NULL, &procname, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	EXCEPINFO info;
	UINT argErr;
	args->rgdispidNamedArgs = &dispid;
	HRESULT hr = disp->Invoke(dispid, IID_NULL, NULL, DISPATCH_METHOD, args, results, &info, &argErr);
	return hr;

	
}

STDMETHODIMP CSimpleScriptSite::AddScript(const WCHAR* source)
{
    if (!source)
        return E_POINTER;


	VARIANT vRes;
	HRESULT hr = spJScriptParse->ParseScriptText(source, nullptr, nullptr, nullptr, 0, 0, NULL, &vRes, nullptr);
	return hr;

	
}


