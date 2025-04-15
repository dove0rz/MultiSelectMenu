#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shlobj.h>
#include <shellapi.h>
#include <vector>
#include <string>
#include <strsafe.h>
#include <Shlwapi.h>
#pragma comment (lib, "shlwapi.lib") // only for PathRemove* api
#include <fstream>

// register:   regsvr32.exe MultiSelectMenu.dll
// unregister: regsvr32.exe /u MultiSelectMenu.dll

const std::wstring ExtName = L"MultiSelectExtension";
const CLSID ExtCLSID = { 0x94879487, 0x5987, 0x5987, {0x12, 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xF0} };
HMODULE selfModule = NULL;

// 2 simple string functions

// output        => the output vector
// str           => the input string
// Delimiters    => the token delimiter string to split our input
// QuotationMark => 2 bytes of char, begin/end mark to ignore delimiters.
//                  for example> split2vector(output, "aa (b b) CC", " ", "()"); // output will be {"aa","b b","CC"}
size_t split2vector(std::vector<std::string>* output, std::string str, std::string delimiter, const char* QuotationMark = NULL) {
	size_t i = 0;
	//int depth = 0;
	int prev_pos = 0;
	while (i < str.length()) {
		if (strncmp(&str[i], delimiter.c_str(), delimiter.length()) == 0) { // if found
			int len = i - prev_pos;
			output->push_back(str.substr(prev_pos, len));
			i += delimiter.length(); // 跳過 delimiter 本身
			prev_pos = i;
		}
		if (QuotationMark && str[i] == QuotationMark[0]) { // first meet
			// try to find the end of another QuotationMark
			int skipBytes = 1;
			while (skipBytes < str.length() - i) {
				if (str[i + skipBytes] == QuotationMark[1] && str[i + skipBytes - 1] != '\\') break;
				skipBytes++;
			}
			i += skipBytes + 1;
		}
		else i++;
	}
	output->push_back(str.substr(prev_pos, i - prev_pos));
	return output->size();
}

// fast forward to tokenCharacters, also handle QuotationMarks
// !! quotationMark is 2 bytes !!
// example> "" or () or [] or {}
uint64_t strFindForwardUntil(const char* buffer, int64_t searchSize, std::string tokenCharacters, const char* QuotationMark = NULL) {
	uint64_t idx = 0;
	for (; idx < searchSize; idx++) {
		// handle quotation mark first
		if (QuotationMark != NULL && buffer[idx] == QuotationMark[0]) {
			uint64_t skipBytes = 1;
			while (skipBytes < searchSize - idx) {
				if (buffer[idx + skipBytes] == QuotationMark[1] && buffer[idx + skipBytes - 1] != '\\') break;
				skipBytes++;
			}
			idx += skipBytes + 1;
		}
		// check if we meet tokens
		char ch = buffer[idx];
		for (int i = 0; i < tokenCharacters.size(); i++) {
			if (ch == tokenCharacters[i]) return idx;
		}
	}
	return searchSize;
}


struct MenuEntry {
	int numSelected; // 0 means any
	std::string fileExtension; // * means any type of extension
	std::string title; // the text you want to show on menu
	std::string cmd, cmdParams;
};

class CShellExt : public IContextMenu, public IShellExtInit {
public:
	LONG m_refCount;
	std::vector<std::string> m_fileList;
	std::vector<MenuEntry> m_menuList;

	CShellExt() {
		m_refCount = 1;
	}

	void readConfig() {
		char strModulePath[MAX_PATH] = { 0 };
		GetModuleFileNameA(selfModule, strModulePath, MAX_PATH);
		//PathRemoveFileSpecA(strModulePath);// remove filename
		PathRemoveExtensionA(strModulePath); // remove extension only
		std::string configName = strModulePath;
		configName += ".conf";

		std::ifstream f(configName.c_str());
		if (!f.is_open()) return;

		m_menuList.clear();
		std::string line;
		while (std::getline(f, line)) {
			// remove left whitespaces
			line.erase(line.begin(), std::find_if(line.begin(), line.end(), [](unsigned char ch) { return !std::isspace(ch); }));
			// remove right whitespaces
			line.erase(std::find_if(line.rbegin(), line.rend(), [](unsigned char ch) { return !std::isspace(ch); }).base(), line.end());
			if (line.empty() || line[0] == '#') continue;

			std::vector<std::string> tokens;
			split2vector(&tokens, line, ";", NULL);

			if (tokens.size() == 4) {
				MenuEntry e;
				if (tokens[0].compare("*") == 0)
					e.numSelected = 0;
				else
					e.numSelected = atoi(tokens[0].c_str());
				e.fileExtension = tokens[1];
				e.title = tokens[2];
				// reserve parameter
				e.cmd = tokens[3]; // write default value first
				e.cmdParams = ""; // write default value first
				uint64_t idx = strFindForwardUntil(tokens[3].c_str(), tokens[3].size(), " ", "\"\"");
				if (idx) {
					e.cmd = tokens[3].substr(0, idx);
					e.cmdParams = tokens[3].substr(idx);
				}
				if (!e.cmd.empty()) { // remove ""
					if (e.cmd.front() == '"' && e.cmd.back() == '"') {
						e.cmd.pop_back(); // remove ending "
						e.cmd.erase(0, 1); // remove beginning "
					}
				}

				if (e.numSelected == 0 || e.numSelected == m_fileList.size())
					m_menuList.push_back(e);
			} else continue;
		}
		f.close();
	}

	// IUnknown

	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) {
		if (!ppv) return E_POINTER;
		if (riid == IID_IUnknown || riid == IID_IContextMenu)
			*ppv = static_cast<IContextMenu*>(this);
		else if (riid == IID_IShellExtInit)
			*ppv = static_cast<IShellExtInit*>(this);
		else {
			*ppv = nullptr;
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}
	
	STDMETHODIMP_(ULONG) AddRef() {
		return  InterlockedIncrement(&m_refCount);
	}

	STDMETHODIMP_(ULONG) Release() {
		ULONG refCount = InterlockedDecrement(&m_refCount);
		if (refCount == 0) {
			delete this;
		}
		return refCount;
	}


	// IContextMenu

	STDMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {
		if (uFlags & CMF_DEFAULTONLY)
			return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

		readConfig();
		
		int numberOfItemsAdded = 0;
		for (UINT i = 0; i < m_menuList.size(); i++) {
			MenuEntry* e = &m_menuList[i];

			if (e->numSelected != 0 && e->numSelected != m_fileList.size()) continue;

			// grab icon from exe
			HICON hIcon = NULL;
			ExtractIconExA(e->cmd.c_str(), 0, NULL, &hIcon, 1); // 取得小圖示
			HBITMAP hBmp = NULL;
			if (hIcon) {
				ICONINFO iconInfo;
				if (GetIconInfo(hIcon, &iconInfo)) {
					hBmp = iconInfo.hbmColor;
					DeleteObject(iconInfo.hbmMask);
				}
			} // hBmp 不需要你釋放，它由 GDI 自己管理，Shell 會釋放圖示用於選單的 hbmpItem

			MENUITEMINFOA mii;
			mii.cbSize = sizeof(MENUITEMINFOA);
			mii.fMask = MIIM_ID | MIIM_STRING | MIIM_BITMAP;
			mii.wID = idCmdFirst + numberOfItemsAdded;
			mii.dwTypeData = (LPSTR)(e->title.c_str());
			mii.hbmpItem = hBmp;
			InsertMenuItemA(hMenu, indexMenu + i, TRUE, &mii);
			numberOfItemsAdded++;
		}
		return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, numberOfItemsAdded);
	}

	STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO lpcmi) {
		if (HIWORD(lpcmi->lpVerb)) return E_FAIL;
		int cmdIndex = LOWORD(lpcmi->lpVerb);

		// build all filepath into one parameter
		std::string command = "\"" + m_menuList[cmdIndex].cmd + "\"";
		command += m_menuList[cmdIndex].cmdParams;
		for (UINT i = 0; i < m_fileList.size(); i++) {
			command += " ";
			command += "\"" + m_fileList[i] + "\"";
		}
		STARTUPINFOA si = { sizeof(si) };
		PROCESS_INFORMATION pi;
		if (!CreateProcessA(NULL, (LPSTR)command.c_str(), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
			return HRESULT_FROM_WIN32(GetLastError());

		CloseHandle(pi.hProcess);
		CloseHandle(pi.hThread);
		return S_OK;
	}

	STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, LPSTR pszName, UINT cchMax) {
		return S_OK;
	}


	// IShellExtInit

	STDMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, IDataObject* pDataObj, HKEY hProgID) {
		if (!pDataObj) return E_INVALIDARG;
		FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM stg = {};
		if (FAILED(pDataObj->GetData(&fe, &stg))) return E_FAIL;
		HDROP hDrop = (HDROP)GlobalLock(stg.hGlobal);
		if (hDrop == NULL) {
			ReleaseStgMedium(&stg);
			return E_FAIL;
		}
		UINT fileCount = DragQueryFileA(hDrop, 0xFFFFFFFF, NULL, 0);
		m_fileList.clear();
		char tmp[1024];
		for (UINT i = 0; i < fileCount; i++) {
			DragQueryFileA(hDrop, i, tmp, sizeof(tmp));
			m_fileList.push_back(tmp);
		}
		GlobalUnlock(stg.hGlobal);
		ReleaseStgMedium(&stg);
		return S_OK;
	}
};

class CShellExtFactory : public IClassFactory {
public:
	LONG m_refCount;
	
	CShellExtFactory() : m_refCount(1) {}

	HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) {
		if (!ppv) return E_POINTER;
		if (riid == IID_IUnknown || riid == IID_IClassFactory)
			*ppv = static_cast<IClassFactory*>(this);
		else {
			*ppv = nullptr;
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}

	ULONG __stdcall AddRef() {
		return InterlockedIncrement(&m_refCount);
	}

	ULONG __stdcall Release() {
		LONG count = InterlockedDecrement(&m_refCount);
		if (count == 0)
			delete this;
		return count;
	}

	HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) {
		if (pUnkOuter != nullptr)
			return CLASS_E_NOAGGREGATION;
		CShellExt* pExt = new CShellExt();
		if (!pExt) return E_OUTOFMEMORY;
		HRESULT hr = pExt->QueryInterface(riid, ppv);
		pExt->Release();
		return hr;
	}
	HRESULT __stdcall LockServer(BOOL fLock) {
		return S_OK;
	}
};


// export table

#pragma comment(linker, "/export:DllGetClassObject=DllGetClassObject,@1")
#pragma comment(linker, "/export:DllCanUnloadNow=DllCanUnloadNow,@2")
#pragma comment(linker, "/export:DllRegisterServer=DllRegisterServer,@3")
#pragma comment(linker, "/export:DllUnregisterServer=DllUnregisterServer,@4")

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
	if (!IsEqualCLSID(rclsid, ExtCLSID))
		return CLASS_E_CLASSNOTAVAILABLE;
	CShellExtFactory* pFactory = new CShellExtFactory();
	if (!pFactory) return E_OUTOFMEMORY;
	HRESULT hr = pFactory->QueryInterface(riid, ppv);
	pFactory->Release();
	return hr;
}

STDAPI DllCanUnloadNow(void) {
	return S_OK;
}

STDAPI DllRegisterServer(void)
{
	HMODULE hModule = NULL;
	if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCWSTR)&DllRegisterServer, &hModule))
		return E_FAIL;

	wchar_t strModulePath[MAX_PATH] = { 0 };
	GetModuleFileNameW(hModule, strModulePath, MAX_PATH);

	WCHAR strCLSID[50];
	if (StringFromGUID2(ExtCLSID, strCLSID, ARRAYSIZE(strCLSID)) == 0)
		return E_FAIL;
	WCHAR keyPath[MAX_PATH];
	StringCchPrintfW(keyPath, ARRAYSIZE(keyPath), L"CLSID\\%s", strCLSID); //_snwprintf_s(keyPath, sizeof(keyPath), L"CLSID\\%s", strCLSID);

	HKEY hKey = NULL;
	if (RegCreateKeyExW(HKEY_CLASSES_ROOT, keyPath, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) != ERROR_SUCCESS)
		return E_FAIL;
	RegSetValueExW(hKey, NULL, 0, REG_SZ, (PBYTE)ExtName.c_str(), (ExtName.size() + 1) * sizeof(WCHAR));
	RegCloseKey(hKey);
	StringCchCatW(keyPath, ARRAYSIZE(keyPath), L"\\InprocServer32");
	if (RegCreateKeyExW(HKEY_CLASSES_ROOT, keyPath, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) != ERROR_SUCCESS)
		return E_FAIL;
	RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)strModulePath, (wcslen(strModulePath) + 1) * sizeof(WCHAR));
	std::wstring threadingModel = L"Apartment";
	RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, (const BYTE*)threadingModel.c_str(), (threadingModel.size() + 1) * sizeof(WCHAR));
	RegCloseKey(hKey);
	// todo: parse config and replace '*' to corresponding file types
	StringCchPrintfW(keyPath, ARRAYSIZE(keyPath), L"*\\shellex\\ContextMenuHandlers\\%s", ExtName.c_str());
	if (RegCreateKeyExW(HKEY_CLASSES_ROOT, keyPath, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) != ERROR_SUCCESS)
		return E_FAIL;
	RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)strCLSID, (lstrlenW(strCLSID) + 1) * sizeof(WCHAR));
	RegCloseKey(hKey);
	
	return S_OK;
}

STDAPI DllUnregisterServer(void)
{
	WCHAR strCLSID[50] = { 0 };
	if (StringFromGUID2(ExtCLSID, strCLSID, ARRAYSIZE(strCLSID)) == 0)
		return E_FAIL;
	// todo: parse config and replace '*' to corresponding file types
	WCHAR  keyPath[MAX_PATH];
	StringCchPrintfW(keyPath, ARRAYSIZE(keyPath), L"*\\shellex\\ContextMenuHandlers\\%s", ExtName.c_str());
	RegDeleteKeyW(HKEY_CLASSES_ROOT, keyPath);
	StringCchPrintf(keyPath, ARRAYSIZE(keyPath), L"CLSID\\%s", strCLSID);
	RegDeleteKey(HKEY_CLASSES_ROOT, keyPath);
	return S_OK;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	selfModule = hModule;
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH: DisableThreadLibraryCalls(hModule); break;
	case DLL_THREAD_ATTACH: break;
	case DLL_THREAD_DETACH: break;
	case DLL_PROCESS_DETACH: break;
	}
	return TRUE;
}

