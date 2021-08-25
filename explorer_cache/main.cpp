#define STRICT
#define STRICT_TYPED_ITEMIDS

#include <Windows.h>
#include <ShlObj.h>
#include <atlbase.h>
#include <atlalloc.h>
#include <ExDispid.h>



constexpr const wchar_t* g_message_caption = L"explorer_cache.exe";



#define message(macro_defined_message) MessageBoxW(nullptr, macro_defined_message, g_message_caption, MB_OK)

#define checkret(macro_defined_argument) if(HRESULT macro_defined_result = (macro_defined_argument); macro_defined_result != S_OK) return macro_defined_result;

#define checkexit(macro_defined_argument, macro_defined_message) if(HRESULT macro_defined_result = (macro_defined_argument); macro_defined_result != S_OK) { message(macro_defined_message); g.destroy(); exit(1); }



struct global_data
{
	struct shell_event_handler : public DShellWindowsEvents
	{
		CComPtr<IConnectionPoint> m_connection_point;

		uint32_t m_ref_cnt;

		uint32_t m_cookie;


		HRESULT create(IShellWindows* unknown)
		{
			m_ref_cnt = 1;

			m_cookie = 0;

			return Connect(unknown);
		}

		void destroy()
		{
			Disconnect();
		}

		/* IUnknown */
		IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
		{
			if (riid == IID_IUnknown || riid == IID_IDispatch || riid == DIID_DShellWindowsEvents)
			{
				*ppv = static_cast<DShellWindowsEvents*>(static_cast<IDispatch*>(this));

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
			riid; lcid; wFlags; pexcepinfo; puArgErr; pdispparams;

			if (pvarResult)
				VariantInit(pvarResult);

			if (dispid == DISPID_WINDOWREGISTERED || dispid == DISPID_WINDOWREVOKED)
				g.rebuild_explorer_list();

			return S_OK;
		}

		HRESULT Connect(IUnknown* punk)
		{
			CComPtr<IConnectionPointContainer> connection_point_container;

			checkret(punk->QueryInterface(IID_PPV_ARGS(&connection_point_container)));

			checkret(connection_point_container->FindConnectionPoint(DIID_DShellWindowsEvents, &m_connection_point));

			return m_connection_point->Advise(this, reinterpret_cast<DWORD*>(&m_cookie));
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

	struct browser_event_handler : DWebBrowserEvents
	{
		CComPtr<IConnectionPoint> m_connection_point;

		uint32_t m_ref_cnt = 0;

		uint32_t m_cookie = 0;

		HWND m_window = nullptr;

		CComHeapPtr<ITEMIDLIST_ABSOLUTE> m_location;

		uint32_t m_openclose_id = 0;



		HRESULT create(HWND window, IDispatch* dispatch, uint32_t openclose_id, ITEMIDLIST_ABSOLUTE* location)
		{
			m_ref_cnt = 1;

			m_cookie = 0;

			m_window = window;

			m_openclose_id = openclose_id;

			m_location.Attach(location);

			return Connect(dispatch);
		}

		void destroy()
		{
			m_location.Detach();

			Disconnect();
		}

		void update_openclose_id(uint32_t openclose_id)
		{
			m_openclose_id = openclose_id;
		}

		/* IUnknown */
		IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
		{
			if (riid == IID_IUnknown || riid == IID_IDispatch || riid == DIID_DWebBrowserEvents)
			{
				*ppv = static_cast<DWebBrowserEvents*>(static_cast<IDispatch*>(this));

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
			riid; lcid; wFlags; pexcepinfo; puArgErr; pvarResult;

			if (pvarResult)
				VariantInit(pvarResult);

			if (dispid == DISPID_NAVIGATECOMPLETE)
			{
				m_location.Free();

				m_location.Attach(ILCreateFromPathW(pdispparams->rgvarg[0].bstrVal));
			}
			else if (dispid == DISPID_QUIT)
				g.last_closed_location.Attach(m_location.Detach());

			return S_OK;
		}

		HRESULT Connect(IUnknown* punk)
		{
			CComPtr<IConnectionPointContainer> connection_point_container;

			checkret(punk->QueryInterface(IID_PPV_ARGS(&connection_point_container)));

			checkret(connection_point_container->FindConnectionPoint(DIID_DWebBrowserEvents, &m_connection_point));

			return m_connection_point->Advise(this, reinterpret_cast<DWORD*>(&m_cookie));
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



	CComPtr<IShellWindows> shell_windows{};

	CComHeapPtr<ITEMIDLIST_ABSOLUTE> last_closed_location{};

	shell_event_handler openclose_handler{};

	browser_event_handler navigation_handlers[32]{};

	uint32_t slots_occupied{};

	uint32_t openclose_cnt{};



	HRESULT rebuild_explorer_list()
	{
		CComPtr<IUnknown> unknown_enum;
		checkret(shell_windows->_NewEnum(&unknown_enum));

		++openclose_cnt;

		CComQIPtr<IEnumVARIANT> enum_variant(unknown_enum);

		for (CComVariant var; enum_variant->Next(1, &var, nullptr) == S_OK; var.Clear())
		{
			if (var.vt != VT_DISPATCH)
				continue;

			// get browser info

			CComPtr<IShellBrowser> curr_shell_browser;
			if (IUnknown_QueryService(var.pdispVal, SID_STopLevelBrowser, IID_PPV_ARGS(&curr_shell_browser)) != S_OK)
				continue;

			HWND curr_window;
			if (curr_shell_browser->GetWindow(&curr_window) != S_OK)
				continue;

			CComPtr<IShellView> shell_view;
			if (curr_shell_browser->QueryActiveShellView(&shell_view) != S_OK)
				continue;

			CComQIPtr<IPersistIDList> persist_id_list(shell_view);
			if (!persist_id_list)
				return E_FAIL;

			CComHeapPtr<ITEMIDLIST_ABSOLUTE> curr_location;
			if (persist_id_list->GetIDList(&curr_location) != S_OK)
				continue;

			// refresh/create navigation handler

			int32_t slot = get_slot_by_hwnd(curr_window);

			if (slot != -1)
			{
				navigation_handlers[slot].update_openclose_id(openclose_cnt);

				continue;
			}
			else if (create_slot(curr_window, curr_shell_browser, var.pdispVal, curr_location.Detach()) != S_OK)
			{
				message(L"Limit of 32 simultaneously open explorer windows reached");

				break;
			}
		}

		for (int32_t i = 0; i != 32; ++i)
		{
			if (!(slots_occupied & (1 << i)))
				continue;

			if (navigation_handlers[i].m_openclose_id != openclose_cnt)
				destroy_slot(i);
		}

		return S_OK;
	}

	int32_t get_slot_by_hwnd(HWND window)
	{
		for(int32_t i = 0; i != 32; ++i)
			if ((slots_occupied & (1 << i)) && navigation_handlers[i].m_window == window)
			{
				return i;
			}

		return -1;
	}

	HRESULT create_slot(HWND window, IShellBrowser* shell_browser, IDispatch* dispatch, ITEMIDLIST_ABSOLUTE* location)
	{
		for (int32_t i = 0; i != 32; ++i)
			if (!(slots_occupied & (1 << i)))
			{
				slots_occupied |= 1 << i;

				navigation_handlers[i].create(window, dispatch, openclose_cnt, location);

				if(last_closed_location)
					shell_browser->BrowseObject(last_closed_location, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);

				return S_OK;
			}

		return E_FAIL;
	}

	void destroy_slot(int32_t slot)
	{
		slots_occupied &= ~(1 << slot);

		navigation_handlers[slot].destroy();
	}



	void create()
	{
		checkexit(CoInitialize(nullptr), L"Call to CoInitialize failed");

		checkexit(shell_windows.CoCreateInstance(CLSID_ShellWindows), L"Could not create IShellWindows");

		checkexit(openclose_handler.create(shell_windows), L"Could not create openclose_handler");

		rebuild_explorer_list();
	}

	void destroy()
	{
		openclose_handler.destroy();

		for (int32_t i = 0; i != 32; ++i)
			if (slots_occupied & (1 << i))
				navigation_handlers[i].destroy();

		shell_windows.Release();

		last_closed_location.Free();

		CoUninitialize();

		message(L"Finished cleanup");
	}
} g;



int WINAPI WinMain(HINSTANCE instance, HINSTANCE prev_instance, LPSTR command_line, int32_t command_show)
{
	command_show; command_line; prev_instance; instance;

	g.create();

	MSG message;

	while (GetMessageW(&message, nullptr, 0, 0))
	{
		TranslateMessage(&message);

		DispatchMessageW(&message);
	}

	g.destroy();

	return 0;
}
