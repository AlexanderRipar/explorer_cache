#include <cstdint>

#include <cstdio>

#include <Windows.h>
#include <UIAutomation.h>
#include <exdisp.h>
#include <shobjidl_core.h>
#include <shlobj_core.h>
#include <rpcdce.h>
#include <wrl\client.h>

#pragma comment(lib, "Rpcrt4.lib")

template<typename T>
using comptr = Microsoft::WRL::ComPtr<T>;

void notify_user(const wchar_t* message)
{
	MessageBoxW(nullptr, message, L"Explorer Cache", MB_OK);
}

void flee_(const wchar_t* message) noexcept
{
	notify_user(message);

	exit(1);
}

#define flee(macro_defined_argument, macro_defined_message) if(HRESULT macro_defined_rst = (macro_defined_argument); macro_defined_rst != S_OK) flee_(macro_defined_message);

struct global_data_struct
{
	struct directory_change_event_handler : IUIAutomationPropertyChangedEventHandler
	{
	private:

		LONG ref_cnt = 1;

	public:

		ULONG STDMETHODCALLTYPE AddRef()
		{
			return ++ref_cnt;
		}

		ULONG STDMETHODCALLTYPE Release()
		{
			return --ref_cnt;
		}

		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
		{
			if (riid == __uuidof(IUnknown))
				*ppInterface = static_cast<IUIAutomationPropertyChangedEventHandler*>(this);
			else if (riid == __uuidof(IUIAutomationPropertyChangedEventHandler))
				*ppInterface = static_cast<IUIAutomationPropertyChangedEventHandler*>(this);
			else
			{
				*ppInterface = nullptr;

				return E_NOINTERFACE;
			}
			
			AddRef();
			
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE HandlePropertyChangedEvent(IUIAutomationElement* pSender, PROPERTYID propertyID, VARIANT newValue)
		{
			newValue;

			// Just a security check that nothing went horribly wrong...
			if (propertyID != UIA_NamePropertyId)
			{
				notify_user(L"Unexpected property change ID\n");

				return S_OK;
			}

			// Find the slot corresponding to the changed address element

			uint32_t slot = ~0u;

			for (uint32_t s = 0; s != 32; ++s)
			{
				BOOL equal_elems = FALSE;

				if (global.explorer_address_elements[s].Get() && global.ui_automation->CompareElements(pSender, global.explorer_address_elements[s].Get(), &equal_elems) != S_OK)
				{
					notify_user(L"Element comparison regarding Folder change failed");

					return S_OK;
				}

				if (equal_elems)
					slot = s;
			}

			if (slot == ~0)
			{
				notify_user(L"Could not find Automation Element corresponding to changed Explorer Window");

				return S_OK;
			}

			global.get_explorer_window_address(slot);

			return S_OK;
		}
	};

	struct open_close_event_handler : public IUIAutomationEventHandler
	{
	private:

		LONG ref_cnt = 1;

	public:

		ULONG STDMETHODCALLTYPE AddRef()
		{
			return ++ref_cnt;
		}

		ULONG STDMETHODCALLTYPE Release()
		{
			return --ref_cnt;
		}

		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
		{
			if (riid == __uuidof(IUnknown))
				*ppInterface = static_cast<IUIAutomationEventHandler*>(this);
			else if (riid == __uuidof(IUIAutomationEventHandler))
				*ppInterface = static_cast<IUIAutomationEventHandler*>(this);
			else
			{
				*ppInterface = nullptr;

				return E_NOINTERFACE;
			}

			AddRef();

			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE HandleAutomationEvent(IUIAutomationElement* pSender, EVENTID eventID)
		{
			if (!is_explorer_window(pSender))
				return S_OK;

			switch (eventID)
			{
			case UIA_Window_WindowOpenedEventId:

				pSender->AddRef();

				UIA_HWND hwnd;

				if (pSender->get_CurrentNativeWindowHandle(&hwnd) != S_OK)
				{
					notify_user(L"Could not get HWND of opened explorer window");

					return S_OK;
				}

				global.set_explorer_window_address(hwnd);

				global.add_window(pSender);

				break;



			case UIA_Window_WindowClosedEventId:

				global.remove_window(pSender);

				break;



			default:

				notify_user(L"Unknown event received\n");

				break;
			}

			return S_OK;
		}
	};





	BSTR cabinet_window_class_name{};

	BSTR explorer_window_class_name{};

	comptr<IShellWindows> shell_windows{};

	comptr<IUIAutomation> ui_automation{};

	comptr<IShellFolder> shell_folder{};

	comptr<IUIAutomationCondition> address_id_search_condition{};

	open_close_event_handler open_close_handler{};

	directory_change_event_handler directory_change_handler{};

	uint32_t slots_occupied_bits{};

	comptr<IUIAutomationElement> explorer_window_elements[32]{};

	comptr<IUIAutomationElement> explorer_address_elements[32]{};

	LPITEMIDLIST explorer_addresses[32]{};

	LPITEMIDLIST last_closed_explorer_address{};




	global_data_struct() noexcept
	{
		flee(CoInitializeEx(nullptr, COINIT_MULTITHREADED), 
			L"Could not initialize COM context");

		flee(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_LOCAL_SERVER, IID_IShellWindows, (&shell_windows)),
			L"Could not create ShellWindows instance");

		flee(SHGetDesktopFolder(&shell_folder),
			L"Could not create desktop shell folder");

		flee(CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_IUIAutomation, (&ui_automation)), 
			L"Could not create UIAutomation instance");

		comptr<IUIAutomationElement> desktop_root;

		flee(ui_automation->GetRootElement(&desktop_root), 
			L"Could not acquire root automation element");

		VARIANT address_automation_id_var;
		address_automation_id_var.vt = VT_BSTR;
		address_automation_id_var.bstrVal = SysAllocString(L"1001");;

		flee(ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, address_automation_id_var, &address_id_search_condition), 
			L"Could not create Search Condition to find Explorer Address Bar");

		flee(ui_automation->AddAutomationEventHandler(UIA_Window_WindowOpenedEventId, desktop_root.Get(), TreeScope_Children, nullptr, &open_close_handler), 
			L"Could not register Event Handler for closing of Explorer Windows");

		flee(ui_automation->AddAutomationEventHandler(UIA_Window_WindowClosedEventId, desktop_root.Get(), TreeScope_Children, nullptr, &open_close_handler), 
			L"Could not register Event Handler for opening of Explorer Windows");

		cabinet_window_class_name = SysAllocString(L"CabinetWClass");

		explorer_window_class_name = SysAllocString(L"ExplorerWClass");

		comptr<IUIAutomationCondition> cabinet_condition;
		
		comptr<IUIAutomationCondition> explorer_condition;
		
		comptr<IUIAutomationCondition> search_condition;
		
		VARIANT cabinet_var;
		cabinet_var.vt = VT_BSTR;
		cabinet_var.bstrVal = cabinet_window_class_name;
		
		VARIANT explorer_var;
		explorer_var.vt = VT_BSTR;
		explorer_var.bstrVal = explorer_window_class_name;
		
		flee(ui_automation->CreatePropertyCondition(UIA_ClassNamePropertyId, cabinet_var, &cabinet_condition), 
			L"Could not create CabinetWClass Property Condition");
		
		flee(ui_automation->CreatePropertyCondition(UIA_ClassNamePropertyId, explorer_var, &explorer_condition), 
			L"Could not create ExplorerWClass Property Condition");
		
		flee(ui_automation->CreateOrCondition(cabinet_condition.Get(), explorer_condition.Get(), &search_condition),
			L"Could not create Explorer-Window search condition");
		
		// Perform search
		
		comptr<IUIAutomationElementArray> found;
		
		flee(desktop_root->FindAll(TreeScope_Children, search_condition.Get(), &found),
			L"Initial search for Explorer Windows failed");
		
		int32_t elem_cnt;
		
		flee(found->get_Length(&elem_cnt),
			L"Could not acquire number of initial Explorer Windows");
		
		for (int32_t i = 0; i != elem_cnt; ++i)
		{
			comptr<IUIAutomationElement> curr;
		
			if (found->GetElement(i, &curr) != S_OK)
				notify_user(L"Could not get Explorer Window Automation Element from Search Array");
			else
				add_window(curr);
		}

		SysFreeString(address_automation_id_var.bstrVal);
	}

	~global_data_struct() noexcept
	{
		notify_user(L"Terminating explorer_cache.exe");

		SysFreeString(cabinet_window_class_name);

		SysFreeString(explorer_window_class_name);

		ui_automation->RemoveAllEventHandlers();
	}

	uint32_t allocate_slot() noexcept
	{
		for (uint32_t i = 0; i != 32; ++i)
			if (!(global.slots_occupied_bits & (1 << i)))
			{
				global.slots_occupied_bits |= 1 << i;

				return i;
			}

		return ~0u;
	}

	void set_explorer_window_address(UIA_HWND explorer_window) noexcept
	{
		if (!last_closed_explorer_address)
		{
			return;
		}

		long shell_cnt;

		if (shell_windows->get_Count(&shell_cnt) != S_OK)
		{
			notify_user(L"Could not get number of current shell windows");

			return;
		}

		for (long i = 0; i != shell_cnt; ++i)
		{
			comptr<IDispatch> dispatch;

			VARIANT index_var;
			index_var.vt = VT_I4;
			index_var.lVal = i;

			if (shell_windows->Item(index_var, &dispatch) != S_OK)
			{
				notify_user(L"Could not access shell window from collection");

				return;
			}

			comptr<IWebBrowserApp> web_browser_app;

			if (dispatch->QueryInterface(IID_PPV_ARGS(&web_browser_app)) != S_OK)
			{
				notify_user(L"Could not get WebBrowserApp interface from shell window");

				return;
			}

			SHANDLE_PTR hwnd;

			if (web_browser_app->get_HWND(&hwnd) != S_OK)
			{
				notify_user(L"Could not get HWND from WebBrowserApp");

				return;
			}

			if (static_cast<uint64_t>(hwnd) == reinterpret_cast<uint64_t>(explorer_window))
			{
				comptr<IServiceProvider> service_provider;

				if (web_browser_app->QueryInterface(IID_PPV_ARGS(&service_provider)) != S_OK)
				{
					notify_user(L"Could not get Service Provider from WebBrowserApp");

					return;
				}

				comptr<IShellBrowser> shell_browser;

				if (service_provider->QueryService(SID_STopLevelBrowser, shell_browser.GetAddressOf()) != S_OK)
				{
					notify_user(L"Could not get top level shell browser from service provider");

					return;
				}

				if (shell_browser->BrowseObject(last_closed_explorer_address, SBSP_SAMEBROWSER | SBSP_ABSOLUTE) != S_OK)
					notify_user(L"Could not browse opened Explorer window to target folder");

				return;
			}
		}

		notify_user(L"Could not explorer window to change path");
	}

	void add_window(comptr<IUIAutomationElement> explorer_window) noexcept
	{
		uint32_t slot = allocate_slot();

		if (slot == ~0u)
		{
			notify_user(L"Reached maximum number of manageable Explorer Windows (32). Further windows are ignored.");

			return;
		}

		explorer_window_elements[slot] = explorer_window;

		explorer_address_elements[slot] = nullptr;

		comptr<IUIAutomationElement> address_bar;

		if (explorer_window->FindFirst(TreeScope_Descendants, address_id_search_condition.Get(), &address_bar) != S_OK || !address_bar)
			return; // This is relevant if the window corresponding to explorer_window is minimized

		PROPERTYID name_property_ids[]{ UIA_NamePropertyId };

		if (ui_automation->AddPropertyChangedEventHandlerNativeArray(address_bar.Get(), TreeScope_Element, nullptr, &directory_change_handler, name_property_ids, _countof(name_property_ids)) != S_OK)
			notify_user(L"Could not add event handler for created Explorer Window\n");
		else
		{
			explorer_address_elements[slot] = address_bar;

			get_explorer_window_address(slot);
		}
	}

	void remove_window(comptr<IUIAutomationElement> explorer_window) noexcept
	{
		for (uint32_t i = 0; i != 32; ++i)
		{
			BOOL element_found = 0;

			if(explorer_window_elements[i].Get())
				flee(ui_automation->CompareElements(explorer_window.Get(), explorer_window_elements[i].Get(), &element_found), 
					L"Element comparison on removed Explorer Window Element failed. This is a critical failure leading to a potential memory leak. Shutting down");

			if (element_found)
			{
				if (explorer_address_elements[i].Get())
				{
					flee(ui_automation->RemovePropertyChangedEventHandler(explorer_address_elements[i].Get(), &directory_change_handler),
						L"Failed to remove Event Handler for a closed Explorer Window. This is a critical failure leading to a potential memory leak. Shutting down");

					&explorer_address_elements[i];

					&explorer_window_elements[i];

					slots_occupied_bits &= ~(1u << i);

					last_closed_explorer_address = explorer_addresses[i];

					explorer_addresses[i] = nullptr;

				}
				else
				{
					&explorer_window_elements[i];

					slots_occupied_bits &= (1u << i);
				}

				return;
			}
		}

		flee(1, L"Could not locate closed explorer window. This is a critical failure leading to a potential memory leak. Shutting down");
	}

	static bool is_explorer_window(comptr<IUIAutomationElement> elem) noexcept
	{
		wchar_t* class_name;

		if (HRESULT rst = elem->get_CurrentClassName(&class_name); rst != S_OK)
			return false;

		bool is_cabinet = true;

		for (uint32_t i = 0; global.cabinet_window_class_name[i]; ++i)
			if (class_name[i] != global.cabinet_window_class_name[i])
			{
				is_cabinet = false;
				break;
			}

		bool is_explorer = true;

		for (uint32_t i = 0; global.explorer_window_class_name[i]; ++i)
			if (class_name[i] != global.explorer_window_class_name[i])
			{
				is_explorer = false;
				break;
			}

		return is_cabinet || is_explorer;
	}

	void get_explorer_window_address(uint32_t slot) noexcept
	{
		// Get hwnd of explorer window element corresponding to address element

		UIA_HWND changed_hwnd;

		if (global.explorer_window_elements[slot]->get_CurrentNativeWindowHandle(&changed_hwnd) != S_OK)
		{
			notify_user(L"Could not get HWND for changed Explorer Window");

			return;
		}

		// Enumerate open shell windows

		long shell_cnt;

		if (global.shell_windows->get_Count(&shell_cnt) != S_OK)
		{
			notify_user(L"Could not get number of current shell windows");

			return;
		}

		// Check which Shell Window corresponds to the changed address element

		for (long i = 0; i != shell_cnt; ++i)
		{
			comptr<IDispatch> dispatch;

			VARIANT index_var;
			index_var.vt = VT_I4;
			index_var.lVal = i;

			if (global.shell_windows->Item(index_var, &dispatch) != S_OK)
			{
				notify_user(L"Could not access shell window from collection");

				return;
			}

			comptr<IWebBrowserApp> web_browser_app;

			if (dispatch->QueryInterface(IID_PPV_ARGS(&web_browser_app)) != S_OK)
			{
				notify_user(L"Could not get WebBrowserApp interface from shell window");

				return;
			}

			SHANDLE_PTR shell_hwnd;

			if (web_browser_app->get_HWND(&shell_hwnd) != S_OK)
			{
				notify_user(L"Could not get HWND from WebBrowserApp");

				return;
			}

			// Check if we found the changed window

			if (reinterpret_cast<void*>(shell_hwnd) == changed_hwnd)
			{
				comptr<IServiceProvider> service_provider;

				if (web_browser_app->QueryInterface(service_provider.GetAddressOf()) != S_OK)
				{
					notify_user(L"Could not acquire ServiceProvider from WebBrowserApp");

					return;
				}

				comptr<IShellBrowser> shell_browser;

				if (service_provider->QueryService(SID_STopLevelBrowser, shell_browser.GetAddressOf()) != S_OK)
				{
					notify_user(L"Could not acquire ShellBrowser from Service Provider");

					return;
				}

				comptr<IShellView> shell_view;

				if (shell_browser->QueryActiveShellView(shell_view.GetAddressOf()) != S_OK)
				{
					notify_user(L"Could not get ShellView from ShellBrowser");

					return;
				}

				comptr<IFolderView> folder_view;

				if (shell_view->QueryInterface(folder_view.GetAddressOf()) != S_OK)
				{
					notify_user(L"Could not get FolderView from ShellView");

					return;
				}

				comptr<IPersistFolder2> persist_folder;

				if (folder_view->GetFolder(IID_IPersistFolder2, reinterpret_cast<void**>(persist_folder.GetAddressOf())) != S_OK)
				{
					notify_user(L"Could not get PersistFolder2 from FolderView");

					return;
				}

				// Get new address as LPITEMIDLIST

				CoTaskMemFree(global.explorer_addresses[slot]);

				if (persist_folder->GetCurFolder(global.explorer_addresses + slot) != S_OK)
				{
					notify_user(L"Could not get new Folder of changed File Explorer Window");

					return;
				}

				return;
			}
		}

		notify_user(L"Could not find Explorer window related to address change");
	}

} global{};

int APIENTRY WinMain(HINSTANCE instance, HINSTANCE prev_instance, LPSTR command_line, int show_commands)
{
	instance;
	prev_instance;
	command_line;
	show_commands;

	SleepEx(INFINITE, 1);
}
