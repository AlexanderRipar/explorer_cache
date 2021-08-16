//#define UNICODE
//#define _UNICODE
#define STRICT
#define STRICT_TYPED_ITEMIDS

#include <Windows.h>
#include <Ole2.h>

#include <ShlObj.h>
#include <atlbase.h>
#include <atlalloc.h>
#include <ExDisp.h>
#include <ExDispid.h>

#include <cstdio>

constexpr const wchar_t* g_message_caption = L"explorer_cache.exe";

CComPtr<IShellWindows> g_shell_windows;





#define checkret(macro_defined_argument) if(HRESULT macro_defined_result = (macro_defined_argument); macro_defined_result != S_OK) return macro_defined_result;

#define checkexit(macro_defined_argument, macro_defined_message) if(HRESULT macro_defined_result = (macro_defined_argument); macro_defined_result != S_OK) { MessageBoxW(nullptr, macro_defined_message, g_message_caption, MB_OK); exit(1); }



template<typename DispInterface>
struct CDispInterfaceBase : public DispInterface
{

private:

	LONG m_ref_cnt;

	CComPtr<IConnectionPoint> m_connection_point;

	DWORD m_cookie;

public:

	CDispInterfaceBase() : m_ref_cnt(1), m_cookie(0) { }

	// Derived class must implement SimpleInvoke
	virtual HRESULT SimpleInvoke(DISPID dispid, DISPPARAMS* pdispparams, VARIANT* pvarResult) = 0;

	/* IUnknown */
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		if (riid == IID_IUnknown || riid == IID_IDispatch || riid == __uuidof(DispInterface))
		{
			*ppv = static_cast<DispInterface*>(static_cast<IDispatch*>(this));

			AddRef();

			return S_OK;
		}
		else
		{
			*ppv = nullptr;

			return E_NOINTERFACE;
		}
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&m_ref_cnt);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		LONG cRef = InterlockedDecrement(&m_ref_cnt);

		if (cRef == 0) delete this;

		return cRef;
	}

	// *** IDispatch ***
	IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo)
	{
		*pctinfo = 0;

		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
	{
		*ppTInfo = nullptr;

		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		return E_NOTIMPL;
	}

	IFACEMETHODIMP Invoke(DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		if (pvarResult) 
			VariantInit(pvarResult);

		return SimpleInvoke(dispid, pdispparams, pvarResult);
	}

	HRESULT Connect(IUnknown* punk)
	{
		CComPtr<IConnectionPointContainer> connection_point_container;

		checkret(punk->QueryInterface(IID_PPV_ARGS(&connection_point_container)));

		checkret(connection_point_container->FindConnectionPoint(__uuidof(DispInterface), &m_connection_point));

		return m_connection_point->Advise(this, &m_cookie);
	}

	void Disconnect()
	{
		if (m_cookie)
		{
			m_connection_point->Unadvise(m_cookie);

			m_connection_point.Release();

			m_cookie = 0;
		}
	}
};

struct browser_event_handler : public DWebBrowserEvents
{
private:

	LONG m_ref_cnt;

	CComPtr<IConnectionPoint> m_connection_point;

	DWORD m_cookie;

	HWND m_explorer_window;

public:

	browser_event_handler(HWND explorer_window) : m_ref_cnt{ 1 }, m_cookie{ 0 }, m_explorer_window{ explorer_window } { }

	/* IUnknown */
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		*ppv = nullptr;

		HRESULT hr = E_NOINTERFACE;

		if (riid == IID_IUnknown || riid == IID_IDispatch || riid == DIID_DWebBrowserEvents)
		{
			*ppv = static_cast<DWebBrowserEvents*>(static_cast<IDispatch*>(this));

			AddRef();

			hr = S_OK;
		}

		return hr;
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&m_ref_cnt);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		return InterlockedDecrement(&m_ref_cnt);
	}

	// *** IDispatch ***
	IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo)
	{
		*pctinfo = 0;

		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
	{
		*ppTInfo = nullptr;

		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		return E_NOTIMPL;
	}

	IFACEMETHODIMP Invoke(DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		if (pvarResult)
			VariantInit(pvarResult);

		switch (dispid)
		{
		case DISPID_NAVIGATECOMPLETE:

			wprintf(L"%s\n", pdispparams->rgvarg[0].bstrVal);

			break;

		case DISPID_QUIT:

			wprintf(L"Quit\n");

			break;
		
		default:

			wprintf(L"Other (%d)\n", dispid);

			break;
		}

		return S_OK;
	}

	HRESULT Connect(IUnknown* unknown)
	{
		CComPtr<IConnectionPointContainer> connection_point_container;

		checkret(unknown->QueryInterface(IID_PPV_ARGS(&connection_point_container)));

		checkret(connection_point_container->FindConnectionPoint(DIID_DWebBrowserEvents, &m_connection_point));

		return m_connection_point->Advise(this, &m_cookie);
	}

	void Disconnect()
	{
		if (m_cookie)
		{
			m_connection_point->Unadvise(m_cookie);

			m_connection_point.Release();

			m_cookie = 0;
		}
	}
};

struct co_init
{
	co_init() noexcept { checkexit(CoInitialize(nullptr), L"Could not Initialize COM"); }

	~co_init() noexcept { CoUninitialize(); }
};



HRESULT get_location_from_view(IShellBrowser* shell_browser, PWSTR* out_location) noexcept
{
	*out_location = nullptr;

	CComPtr<IShellView> shell_view;
	checkret(shell_browser->QueryActiveShellView(&shell_view));

	CComQIPtr<IPersistIDList> persist_id_list(shell_view);
	if (!persist_id_list)
		return E_FAIL;

	CComHeapPtr<ITEMIDLIST_ABSOLUTE> id_list;
	checkret(persist_id_list->GetIDList(&id_list));

	CComPtr<IShellItem> shell_item;
	checkret(SHCreateItemFromIDList(id_list, IID_PPV_ARGS(&shell_item)));

	return shell_item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, out_location);
}

HRESULT get_browser_info(IUnknown* unknown, HWND* out_window, PWSTR* out_location) noexcept
{
	CComPtr<IShellBrowser> shell_browser;
	checkret(IUnknown_QueryService(unknown, SID_STopLevelBrowser, IID_PPV_ARGS(&shell_browser)));

	checkret(shell_browser->GetWindow(out_window));

	if (get_location_from_view(shell_browser, out_location) == S_OK)
		return S_OK;
}

HRESULT init_windows() noexcept
{
	CComPtr<IUnknown> unknown_enum;
	checkret(g_shell_windows->_NewEnum(&unknown_enum));

	CComQIPtr<IEnumVARIANT> enum_variant(unknown_enum);

	for (CComVariant var; enum_variant->Next(1, &var, nullptr) == S_OK; var.Clear())
	{
		if (var.vt != VT_DISPATCH) continue;

		HWND curr_window;
		CComHeapPtr<WCHAR> curr_location;

		if (get_browser_info(var.pdispVal, &curr_window, &curr_location) != S_OK)
			continue;

		wprintf(L"%p: %s\n", curr_window, curr_location.m_pData);
	}

	return S_OK;
}




int main()
{
	co_init init;

	checkexit(g_shell_windows.CoCreateInstance(CLSID_ShellWindows), L"Could not Create ShellWindows instance");

	checkexit(init_windows(), L"Failed");

	SleepEx(5000, 1);

	g_shell_windows.Release();
}
