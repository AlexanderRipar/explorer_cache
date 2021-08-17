//#define UNICODE
//#define _UNICODE
#define STRICT
#define STRICT_TYPED_ITEMIDS

#include <Windows.h>
#include <windowsx.h>
#include <Ole2.h>
#include <CommCtrl.h>
#include <Shlwapi.h>

#include <ShlObj.h>
#include <atlbase.h>
#include <atlalloc.h>
#include <ExDisp.h>
#include <ExDispid.h>

#include <cstdio>



HRESULT build_explorer_list();



constexpr const wchar_t* g_message_caption = L"explorer_cache.exe";

CComPtr<IShellWindows> g_shell_windows;

HWND g_window_child;

HINSTANCE g_instance;

uint32_t g_openclose_cnt;




void on_size(HWND window, uint32_t state, int32_t cx, int32_t cy)
{
	window, state;

	if (g_window_child)
		MoveWindow(g_window_child, 0, 0, cx, cy, true);
}

BOOL on_create(HWND window, LPCREATESTRUCT create_struct);

void on_destroy(HWND window);

void paint_content(HWND window, PAINTSTRUCT* ps)
{
	window, ps;
}

void on_paint(HWND window)
{
	PAINTSTRUCT ps;
	
	BeginPaint(window, &ps);

	paint_content(window, &ps);

	EndPaint(window, &ps);
}

void on_print_client(HWND window, HDC hdc)
{
	PAINTSTRUCT ps;

	ps.hdc = hdc;

	GetClientRect(window, &ps.rcPaint);

	paint_content(window, &ps);
}

LRESULT on_notify(HWND window, int32_t id_from, NMHDR* nmhdr);

LRESULT CALLBACK window_fn(HWND window, uint32_t message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		HANDLE_MSG(window, WM_CREATE, on_create);

		HANDLE_MSG(window, WM_SIZE, on_size);

		HANDLE_MSG(window, WM_DESTROY, on_destroy);

		HANDLE_MSG(window, WM_PAINT, on_paint);

		HANDLE_MSG(window, WM_NOTIFY, on_notify);

	case WM_PRINTCLIENT:

		on_print_client(window, reinterpret_cast<HDC>(wParam));

		return 0;
	}

	return DefWindowProcW(window, message, wParam, lParam);
}

bool init()
{
	WNDCLASS wc;
	wc.style = 0;
	wc.lpfnWndProc = window_fn;
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = g_instance;
	wc.hIcon = nullptr;
	wc.hCursor = LoadCursorW(g_instance, IDC_ARROW);
	wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
	wc.lpszMenuName = nullptr;
	wc.lpszClassName = L"och_scratch";
	
	if (!RegisterClassW(&wc))
		return false;

	return true;
}



#define checkret(macro_defined_argument) if(HRESULT macro_defined_result = (macro_defined_argument); macro_defined_result != S_OK) return macro_defined_result;

#define checkexit(macro_defined_argument, macro_defined_message) if(HRESULT macro_defined_result = (macro_defined_argument); macro_defined_result != S_OK) { MessageBoxW(nullptr, macro_defined_message, g_message_caption, MB_OK); exit(1); }



template<typename DispInterface>
struct dispatch_interface_base : public DispInterface
{

private:

	LONG m_ref_cnt;

	CComPtr<IConnectionPoint> m_connection_point;

	DWORD m_cookie;

public:

	dispatch_interface_base() : m_ref_cnt(1), m_cookie(0) { }

	virtual ~dispatch_interface_base() {};

	// Derived class must implement SimpleInvoke
	virtual HRESULT simple_invoke(DISPID id, DISPPARAMS* params, VARIANT* var_results) = 0;

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
		iTInfo; lcid;

		*ppTInfo = nullptr;

		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
	{
		riid; rgszNames; cNames; lcid; rgDispId;

		return E_NOTIMPL;
	}

	IFACEMETHODIMP Invoke(DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		riid; lcid; wFlags; pexcepinfo; puArgErr;

		if (pvarResult) 
			VariantInit(pvarResult);

		return simple_invoke(dispid, pdispparams, pvarResult);
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



void update_text(HWND window, PCWSTR text);

struct browser_event_sink : public dispatch_interface_base<DWebBrowserEvents>
{
private:

	HWND m_window;

public:

	browser_event_sink(HWND window) : m_window{ window } {}

	IFACEMETHODIMP simple_invoke(DISPID id, DISPPARAMS* params, VARIANT* var_results)
	{
		var_results;

		switch (id)
		{
		case DISPID_NAVIGATECOMPLETE:
			update_text(m_window, params->rgvarg[0].bstrVal);
			break;

		case DISPID_QUIT:
			update_text(m_window, L"<closed>");
			break;
		}

		return S_OK;
	}
};

struct shell_event_sink : public dispatch_interface_base<DShellWindowsEvents>
{
	IFACEMETHODIMP simple_invoke(DISPID id, DISPPARAMS* params, VARIANT* var_results)
	{
		params; var_results;

		if (id == DISPID_WINDOWREGISTERED || id == DISPID_WINDOWREVOKED)
			build_explorer_list();

		return S_OK;
	}

	~shell_event_sink() {};
};

CComPtr<shell_event_sink> g_shell_handler;

struct item_info
{
	HWND m_window;

	CComPtr<browser_event_sink> handler;

	uint32_t last_openclose_cnt;

	item_info(HWND window, IDispatch* dispatch) : m_window{ window }, last_openclose_cnt{ g_openclose_cnt }
	{
		handler.Attach(new(std::nothrow) browser_event_sink(window));

		if (handler)
			handler->Connect(dispatch);
	}

	~item_info() 
	{
		if (handler)
			handler->Disconnect();
	}
};



LRESULT on_notify(HWND window, int32_t id_from, NMHDR* nmhdr)
{
	window;

	if (id_from == 1)
	{
		if (nmhdr->code == LVN_DELETEITEM)
		{
			NMLISTVIEW* nm_listview = CONTAINING_RECORD(nmhdr, NMLISTVIEW, hdr);

			delete reinterpret_cast<item_info*>(nm_listview->lParam);
		}
	}

	return 0;
}

BOOL on_create(HWND window, LPCREATESTRUCT create_struct)
{
	create_struct;

	g_window_child = CreateWindow(WC_LISTVIEW, nullptr, LVS_LIST | WS_CHILD | WS_VISIBLE | WS_HSCROLL | WS_VSCROLL, 0, 0, 0, 0, window, reinterpret_cast<HMENU>(1), g_instance, 0);

	if (g_shell_windows.CoCreateInstance(CLSID_ShellWindows) != S_OK)
		return FALSE;

	build_explorer_list();

	g_shell_handler.Attach(new(std::nothrow) shell_event_sink);

	g_shell_handler->Connect(g_shell_windows);

	return TRUE;
}

void on_destroy(HWND window)
{
	window;

	g_shell_windows.Release();

	if (g_shell_handler)
	{
		g_shell_handler->Disconnect();

		g_shell_handler.Release();
	}

	PostQuitMessage(0);
}



item_info* get_item_by_index(int32_t index)
{
	LVITEM item{};
	item.mask = LVIF_PARAM;
	item.iItem = index;
	
	ListView_GetItem(g_window_child, &item);

	return reinterpret_cast<item_info*>(item.lParam);
}

item_info* get_item_by_window(HWND window, int32_t* out_index)
{
	int32_t index = ListView_GetItemCount(g_window_child);

	while (--index >= 0)
	{
		item_info* info = get_item_by_index(index);

		if (info->m_window == window)
		{
			if (out_index)
				*out_index = index;

			return info;
		}
	}

	return nullptr;
}

void update_text(HWND window, PCWSTR text)
{
	int index;

	if (get_item_by_window(window, &index))
		ListView_SetItemText(g_window_child, index, 0, const_cast<PWSTR>(text));
}

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

	return get_location_from_view(shell_browser, out_location);
}

HRESULT build_explorer_list()
{
	CComPtr<IUnknown> unknown_enum;
	checkret(g_shell_windows->_NewEnum(&unknown_enum));

	++g_openclose_cnt;

	CComQIPtr<IEnumVARIANT> enum_variant(unknown_enum);

	for (CComVariant var; enum_variant->Next(1, &var, nullptr) == S_OK; var.Clear())
	{
		if (var.vt != VT_DISPATCH)
			continue;

		HWND curr_window;

		CComHeapPtr<WCHAR> curr_location;

		if (get_browser_info(var.pdispVal, &curr_window, &curr_location) != S_OK)
			continue;

		item_info* info = get_item_by_window(curr_window, nullptr);

		if (info)
		{
			info->last_openclose_cnt = g_openclose_cnt;

			continue;
		}
		else
			info = new(std::nothrow) item_info(curr_window, var.pdispVal);

		if (!info)
			continue;

		LVITEM item{};
		item.mask = LVIF_TEXT | LVIF_PARAM;
		item.iItem = MAXLONG;
		item.iSubItem = 0;
		item.pszText = curr_location;
		item.lParam = reinterpret_cast<LPARAM>(info);
		
		if (ListView_InsertItem(g_window_child, &item) < 0)
			delete info;
	}

	int32_t item_idx = ListView_GetItemCount(g_window_child);

	while (--item_idx >= 0)
	{
		item_info* info = get_item_by_index(item_idx);

		if (info->last_openclose_cnt != g_openclose_cnt)
			ListView_DeleteItem(g_window_child, item_idx);
	}

	return S_OK;
}

int WINAPI WinMain(HINSTANCE instance, HINSTANCE prev_instance, LPSTR command_line, int32_t command_show)
{
	command_line, prev_instance;

	g_instance = instance;

	if (!init())
		return 0;

	if (CoInitialize(nullptr) != S_OK)
		return 0;

	HWND window = CreateWindowW(L"och_scratch", L"scratch ui", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, nullptr, nullptr, instance, 0);

	ShowWindow(window, command_show);

	MSG message;

	while (GetMessageW(&message, nullptr, 0, 0))
	{
		TranslateMessage(&message);

		DispatchMessageW(&message);
	}

	CoUninitialize();

	return 0;
}
