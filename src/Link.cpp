// Link.cpp: A wrapper for existing hyper links.

#include "stdafx.h"
#include "Link.h"
#include "ClassFactory.h"
#include "CWindowEx2.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP Link::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ILink, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


Link::Properties::~Properties()
{
	if(pOwnerSysLink) {
		pOwnerSysLink->Release();
	}
}

HWND Link::Properties::GetSysLinkHWnd(void)
{
	ATLASSUME(pOwnerSysLink);

	OLE_HANDLE handle = NULL;
	pOwnerSysLink->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void Link::Attach(int linkIndex)
{
	properties.linkIndex = linkIndex;
}

void Link::Detach(void)
{
	properties.linkIndex = -1;
}

void Link::SetOwner(SysLink* pOwner)
{
	if(properties.pOwnerSysLink) {
		properties.pOwnerSysLink->Release();
	}
	properties.pOwnerSysLink = pOwner;
	if(properties.pOwnerSysLink) {
		properties.pOwnerSysLink->AddRef();
	}
}


STDMETHODIMP Link::get_Caret(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_FOCUSED;
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		*pValue = BOOL2VARIANTBOOL(link.state & LIS_FOCUSED);
	}
	return S_OK;
}

STDMETHODIMP Link::get_DrawAsNormalText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;

	if(properties.pOwnerSysLink->IsComctl32Version610OrNewer()) {
		HWND hWndSysLink = properties.GetSysLinkHWnd();
		ATLASSERT(IsWindow(hWndSysLink));

		LITEM link = {0};
		link.iLink = properties.linkIndex;
		link.mask = LIF_ITEMINDEX | LIF_STATE;
		link.stateMask = LIS_DEFAULTCOLORS;
		if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
			*pValue = BOOL2VARIANTBOOL(link.state & LIS_DEFAULTCOLORS);
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP Link::put_DrawAsNormalText(VARIANT_BOOL newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	if(properties.pOwnerSysLink->IsComctl32Version610OrNewer()) {
		HWND hWndSysLink = properties.GetSysLinkHWnd();
		ATLASSERT(IsWindow(hWndSysLink));

		LITEM link = {0};
		link.iLink = properties.linkIndex;
		link.mask = LIF_ITEMINDEX | LIF_STATE;
		link.stateMask = LIS_DEFAULTCOLORS;
		if(newValue != VARIANT_FALSE) {
			link.state = LIS_DEFAULTCOLORS;
		}
		if(SendMessage(hWndSysLink, LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
			CWindow(hWndSysLink).InvalidateRect(NULL);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP Link::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_ENABLED;
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		*pValue = BOOL2VARIANTBOOL(link.state & LIS_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP Link::put_Enabled(VARIANT_BOOL newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_ENABLED;
	if(newValue != VARIANT_FALSE) {
		link.state = LIS_ENABLED;
	}
	if(SendMessage(hWndSysLink, LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Link::get_Hot(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_HOTTRACK;
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		*pValue = BOOL2VARIANTBOOL(link.state & LIS_HOTTRACK);
	}
	return S_OK;
}

STDMETHODIMP Link::put_Hot(VARIANT_BOOL newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_HOTTRACK;
	if(newValue != VARIANT_FALSE) {
		link.state = LIS_HOTTRACK;
	}
	if(SendMessage(hWndSysLink, LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Link::get_ID(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	*pValue = L"";
	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_ITEMID;
	ZeroMemory(link.szID, MAX_LINKID_TEXT * sizeof(TCHAR));
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		*pValue = _bstr_t(link.szID).Detach();
	}
	return S_OK;
}

STDMETHODIMP Link::put_ID(BSTR newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	CComBSTR bstr;
	CWindow(hWndSysLink).GetWindowText(&bstr);
	CString text = bstr;
	CString newID = newValue;
	CString newControlText = TEXT("");

	#ifdef USE_STL
		std::vector<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		for(std::vector<DOCITEM>::iterator it = documentItems.begin(); it != documentItems.end(); it++) {
			if(it->isLink) {
				newControlText += TEXT("<a");
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					if(!newID.IsEmpty()) {
						newControlText += TEXT(" id=\"") + newID + TEXT("\"");
					}
				} else {
					if(!it->ID.IsEmpty()) {
						newControlText += TEXT(" id=\"") + it->ID + TEXT("\"");
					}
				}
				if(!it->URL.IsEmpty()) {
					newControlText += TEXT(" href=\"") + it->URL + TEXT("\"");
				}
				newControlText += TEXT(">") + it->text + TEXT("</a>");
			} else {
				newControlText += it->text;
			}
		}
	#else
		CAtlList<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		POSITION p = documentItems.GetHeadPosition();
		while(p) {
			DOCITEM item = documentItems.GetAt(p);
			if(item.isLink) {
				newControlText += TEXT("<a");
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					if(!newID.IsEmpty()) {
						newControlText += TEXT(" id=\"") + newID + TEXT("\"");
					}
				} else {
					if(!item.ID.IsEmpty()) {
						newControlText += TEXT(" id=\"") + item.ID + TEXT("\"");
					}
				}
				if(!item.URL.IsEmpty()) {
					newControlText += TEXT(" href=\"") + item.URL + TEXT("\"");
				}
				newControlText += TEXT(">") + item.text + TEXT("</a>");
			} else {
				newControlText += item.text;
			}
			documentItems.GetNext(p);
		}
	#endif
	CWindow(hWndSysLink).SetWindowText(newControlText);

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_ITEMID;
	StringCchCopyW(link.szID, MAX_LINKID_TEXT, OLE2W(newValue));
	if(SendMessage(hWndSysLink, LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Link::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.linkIndex;
	return S_OK;
}

STDMETHODIMP Link::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	*pValue = L"";

	CComBSTR bstr;
	CWindow(hWndSysLink).GetWindowText(&bstr);
	CString text = bstr;

	#ifdef USE_STL
		std::vector<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		for(std::vector<DOCITEM>::iterator it = documentItems.begin(); it != documentItems.end(); it++) {
			if(it->isLink) {
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					*pValue = _bstr_t(it->text).Detach();
					break;
				}
			}
		}
	#else
		CAtlList<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		POSITION p = documentItems.GetHeadPosition();
		while(p) {
			DOCITEM item = documentItems.GetAt(p);
			if(item.isLink) {
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					*pValue = _bstr_t(item.text).Detach();
					break;
				}
			}
			documentItems.GetNext(p);
		}
	#endif
	return S_OK;
}

STDMETHODIMP Link::put_Text(BSTR newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	CComBSTR bstr;
	CWindow(hWndSysLink).GetWindowText(&bstr);
	CString text = bstr;
	CString newText = newValue;
	CString newControlText = TEXT("");

	#ifdef USE_STL
		std::vector<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		for(std::vector<DOCITEM>::iterator it = documentItems.begin(); it != documentItems.end(); it++) {
			if(it->isLink) {
				newControlText += TEXT("<a");
				if(!it->ID.IsEmpty()) {
					newControlText += TEXT(" id=\"") + it->ID + TEXT("\"");
				}
				if(!it->URL.IsEmpty()) {
					newControlText += TEXT(" href=\"") + it->URL + TEXT("\"");
				}
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					newControlText += TEXT(">") + newText + TEXT("</a>");
				} else {
					newControlText += TEXT(">") + it->text + TEXT("</a>");
				}
			} else {
				newControlText += it->text;
			}
		}
	#else
		CAtlList<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		POSITION p = documentItems.GetHeadPosition();
		while(p) {
			DOCITEM item = documentItems.GetAt(p);
			if(item.isLink) {
				newControlText += TEXT("<a");
				if(!item.ID.IsEmpty()) {
					newControlText += TEXT(" id=\"") + item.ID + TEXT("\"");
				}
				if(!item.URL.IsEmpty()) {
					newControlText += TEXT(" href=\"") + item.URL + TEXT("\"");
				}
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					newControlText += TEXT(">") + newText + TEXT("</a>");
				} else {
					newControlText += TEXT(">") + item.text + TEXT("</a>");
				}
			} else {
				newControlText += item.text;
			}
			documentItems.GetNext(p);
		}
	#endif
	CWindow(hWndSysLink).SetWindowText(newControlText);
	return S_OK;
}

STDMETHODIMP Link::get_URL(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	*pValue = L"";
	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_URL;
	ZeroMemory(link.szUrl, L_MAX_URL_LENGTH * sizeof(TCHAR));
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		*pValue = _bstr_t(link.szUrl).Detach();
	}
	return S_OK;
}

STDMETHODIMP Link::put_URL(BSTR newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	CComBSTR bstr;
	CWindow(hWndSysLink).GetWindowText(&bstr);
	CString text = bstr;
	CString newURL = newValue;
	CString newControlText = TEXT("");

	#ifdef USE_STL
		std::vector<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		for(std::vector<DOCITEM>::iterator it = documentItems.begin(); it != documentItems.end(); it++) {
			if(it->isLink) {
				newControlText += TEXT("<a");
				if(!it->ID.IsEmpty()) {
					newControlText += TEXT(" id=\"") + it->ID + TEXT("\"");
				}
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					if(!newURL.IsEmpty()) {
						newControlText += TEXT(" href=\"") + newURL + TEXT("\"");
					}
				} else {
					if(!it->URL.IsEmpty()) {
						newControlText += TEXT(" href=\"") + it->URL + TEXT("\"");
					}
				}
				newControlText += TEXT(">") + it->text + TEXT("</a>");
			} else {
				newControlText += it->text;
			}
		}
	#else
		CAtlList<DOCITEM> documentItems;
		ParseSysLinkText(text, documentItems);
		int linkIndex = -1;
		POSITION p = documentItems.GetHeadPosition();
		while(p) {
			DOCITEM item = documentItems.GetAt(p);
			if(item.isLink) {
				newControlText += TEXT("<a");
				if(!item.ID.IsEmpty()) {
					newControlText += TEXT(" id=\"") + item.ID + TEXT("\"");
				}
				linkIndex++;
				if(linkIndex == properties.linkIndex) {
					if(!newURL.IsEmpty()) {
						newControlText += TEXT(" href=\"") + newURL + TEXT("\"");
					}
				} else {
					if(!item.URL.IsEmpty()) {
						newControlText += TEXT(" href=\"") + item.URL + TEXT("\"");
					}
				}
				newControlText += TEXT(">") + item.text + TEXT("</a>");
			} else {
				newControlText += item.text;
			}
			documentItems.GetNext(p);
		}
	#endif
	CWindow(hWndSysLink).SetWindowText(newControlText);

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_URL;
	StringCchCopyW(link.szUrl, L_MAX_URL_LENGTH, OLE2W(newValue));
	if(SendMessage(hWndSysLink, LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP Link::get_Visited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_VISITED;
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		*pValue = BOOL2VARIANTBOOL(link.state & LIS_VISITED);
	}
	return S_OK;
}

STDMETHODIMP Link::put_Visited(VARIANT_BOOL newValue)
{
	if(properties.linkIndex < 0) {
		return E_FAIL;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = properties.linkIndex;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	link.stateMask = LIS_VISITED;
	if(newValue != VARIANT_FALSE) {
		link.state = LIS_VISITED;
	}
	if(SendMessage(hWndSysLink, LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return S_OK;
	}
	return E_FAIL;
}