// SysLink.cpp: Superclasses SysLink.

#include "stdafx.h"
#include "SysLink.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC SysLink::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


SysLink::SysLink()
{
	properties.font.InitializePropertyWatcher(this, DISPID_SL_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_SL_MOUSEICON);

	SIZEL size = {100, 20};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	currentCaretLink = -1;
	linkUnderMouse = -1;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}
}

DWORD SysLink::_GetViewStatus()
{
	switch(properties.backStyle) {
		case bstTransparent:
			return 0;
			break;
		case bstOpaque:
			return VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE;
			break;
	}
	return VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE;
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP SysLink::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ISysLink, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP SysLink::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<SysLink>::Load(pPropertyBag, pErrorLog);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);

		CComBSTR bstr;
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(SUCCEEDED(hr)) {
			if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
				bstr.ArrayToBSTR(propertyValue.parray);
			} else if(propertyValue.vt == VT_BSTR) {
				bstr = propertyValue.bstrVal;
			}
		} else if(hr == E_INVALIDARG) {
			hr = S_OK;
		}
		put_Text(bstr);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP SysLink::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<SysLink>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	if(SUCCEEDED(hr)) {
		VARIANT propertyValue;
		VariantInit(&propertyValue);
		propertyValue.vt = VT_ARRAY | VT_UI1;
		properties.text.BSTRToArray(&propertyValue.parray);
		hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue);
		VariantClear(&propertyValue);
	}
	return hr;
}

STDMETHODIMP SysLink::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 10 VT_I4 properties...
	pSize->QuadPart += 10 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 13 VT_BOOL properties...
	pSize->QuadPart += 13 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 1 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP SysLink::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x11111111/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0100) {
		return E_FAIL;
	}
	LONG subSignature = 0;
	if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}
	if(subSignature != 0x01010101/*4x 0x01 (-> SysLink)*/) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL automaticallyMarkLinksAsVisited = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BackStyleConstants backStyle = static_cast<BackStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL detectDoubleClicks = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HAlignmentConstants hAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL hotTracking = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processReturnKey = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showToolTips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR text;
	if(FAILED(hr = text.ReadFromStream(pStream))) {
		return hr;
	}
	if(!text) {
		text = L"";
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useMnemonic = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useVisualStyle = propertyValue.boolVal;

	BOOL b = flags.dontRecreate;
	flags.dontRecreate = TRUE;
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutomaticallyMarkLinksAsVisited(automaticallyMarkLinksAsVisited);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackStyle(backStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DetectDoubleClicks(detectDoubleClicks);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HAlignment(hAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HotTracking(hotTracking);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessReturnKey(processReturnKey);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowToolTips(showToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseMnemonic(useMnemonic);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseVisualStyle(useVisualStyle);
	if(FAILED(hr)) {
		return hr;
	}
	flags.dontRecreate = b;
	RecreateControlWindow();

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP SysLink::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x11111111/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0100;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x01010101/*4x 0x01 (-> SysLink)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.automaticallyMarkLinksAsVisited);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.hotTracking);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processReturnKey);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showToolTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useMnemonic);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useVisualStyle);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND SysLink::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	if(!RunTimeHelper::IsCommCtrl6()) {
		MessageBox(TEXT("The SysLink control requires at least version 6.0 of comctl32.dll. In order to use it, you have to define a manifest file for your application.\nFor using the control in the VB6 IDE, define a manifest file for VB6.exe."), TEXT("Cannot create control"), MB_ICONERROR | MB_OK);
		return NULL;
	}

	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_LINK_CLASS;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<SysLink>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT SysLink::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("SysLink ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void SysLink::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP SysLink::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<SysLink>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	if(properties.resetTextToName) {
		properties.resetTextToName = FALSE;

		BSTR buffer = SysAllocString(L"");
		if(SUCCEEDED(GetAmbientDisplayName(buffer))) {
			properties.text.AssignBSTR(buffer);
			properties.text = L"<a>" + properties.text + L"</a>";
		} else {
			SysFreeString(buffer);
		}
	}

	return hr;
}

STDMETHODIMP SysLink::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL SysLink::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		if(dialogCode & DLGC_WANTTAB) {
			if(pMessage->wParam == VK_TAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch(pMessage->wParam) {
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
			case VK_HOME:
			case VK_END:
			case VK_NEXT:
			case VK_PRIOR:
				if(dialogCode & DLGC_WANTARROWS) {
					if(!(GetKeyState(VK_SHIFT) & 0x8000) && !(GetKeyState(VK_CONTROL) & 0x8000) && !(GetKeyState(VK_MENU) & 0x8000)) {
						SendMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam);
						hReturnValue = S_OK;
						return TRUE;
					}
				}
				break;
			case VK_TAB:
				if(pMessage->message == WM_KEYDOWN) {
					int numberOfLinks = 0;
					int activeLink = -1;
					LITEM link = {0};
					for(link.iLink = 0; ; link.iLink++) {
						link.mask = LIF_ITEMINDEX | LIF_STATE;
						link.stateMask = LIS_FOCUSED;
						if(SendMessage(LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
							if(link.state & LIS_FOCUSED) {
								activeLink = link.iLink;
							}
						} else {
							numberOfLinks = link.iLink;
							break;
						}
					}
					if(GetKeyState(VK_SHIFT) & 0x8000) {
						// cycle backward
						if(activeLink > 0) {
							// we've not yet reached the first link
							EnterSilentCaretChangeSection();
							link.mask = LIF_ITEMINDEX | LIF_STATE;
							link.stateMask = LIS_FOCUSED;
							link.iLink = activeLink;
							link.state = 0;
							SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
							link.iLink = activeLink - 1;
							link.state = LIS_FOCUSED;
							SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
							hReturnValue = S_OK;
							LeaveSilentCaretChangeSection();

							if(link.iLink != currentCaretLink) {
								Raise_CaretChanged(currentCaretLink, link.iLink);
							}
							return TRUE;
						} else {
							if(-1 != currentCaretLink) {
								Raise_CaretChanged(currentCaretLink, -1);
							}
						}
					} else {
						// cycle forward
						if(numberOfLinks > 0 && activeLink < numberOfLinks - 1) {
							// we've not yet reached the last link
							EnterSilentCaretChangeSection();
							link.mask = LIF_ITEMINDEX | LIF_STATE;
							link.stateMask = LIS_FOCUSED;
							link.iLink = activeLink;
							link.state = 0;
							SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
							link.iLink = activeLink + 1;
							link.state = LIS_FOCUSED;
							SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
							hReturnValue = S_OK;
							LeaveSilentCaretChangeSection();

							if(link.iLink != currentCaretLink) {
								Raise_CaretChanged(currentCaretLink, link.iLink);
							}
							return TRUE;
						} else {
							if(-1 != currentCaretLink) {
								Raise_CaretChanged(currentCaretLink, -1);
							}
						}
					}
				}
				break;
		}
	}
	return CComControl<SysLink>::PreTranslateAccelerator(pMessage, hReturnValue);
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP SysLink::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);

	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP SysLink::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP SysLink::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);

	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP SysLink::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP SysLink::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Accessibility:
			*pName = GetResString(IDPC_ACCESSIBILITY).Detach();
			return S_OK;
			break;
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP SysLink::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_SL_APPEARANCE:
		case DISPID_SL_BACKSTYLE:
		case DISPID_SL_BORDERSTYLE:
		case DISPID_SL_HALIGNMENT:
		case DISPID_SL_MOUSEICON:
		case DISPID_SL_MOUSEPOINTER:
		case DISPID_SL_USEVISUALSTYLE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_SL_AUTOMATICALLYMARKLINKSASVISITED:
		case DISPID_SL_CARETLINK:
		case DISPID_SL_DETECTDOUBLECLICKS:
		case DISPID_SL_DISABLEDEVENTS:
		case DISPID_SL_DONTREDRAW:
		case DISPID_SL_HOTTRACKING:
		case DISPID_SL_HOVERTIME:
		case DISPID_SL_PROCESSCONTEXTMENUKEYS:
		case DISPID_SL_PROCESSRETURNKEY:
		case DISPID_SL_RIGHTTOLEFT:
		case DISPID_SL_SHOWTOOLTIPS:
		case DISPID_SL_USEMNEMONIC:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_SL_BACKCOLOR:
		case DISPID_SL_FORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_SL_APPID:
		case DISPID_SL_APPNAME:
		case DISPID_SL_APPSHORTNAME:
		case DISPID_SL_BUILD:
		case DISPID_SL_CHARSET:
		case DISPID_SL_HWND:
		case DISPID_SL_HWNDTOOLTIP:
		case DISPID_SL_ISRELEASE:
		case DISPID_SL_PROGRAMMER:
		case DISPID_SL_TESTER:
		case DISPID_SL_TEXT:
		case DISPID_SL_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_SL_REGISTERFOROLEDRAGDROP:
		case DISPID_SL_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_SL_FONT:
		case DISPID_SL_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_SL_LINKS:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_SL_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString SysLink::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString SysLink::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString SysLink::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=LabelControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString SysLink::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString SysLink::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString SysLink::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP SysLink::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_SL_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_SL_BACKSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.backStyle), description);
			break;
		case DISPID_SL_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_SL_HALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
			break;
		case DISPID_SL_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_SL_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_SL_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		default:
			return IPerPropertyBrowsingImpl<SysLink>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP SysLink::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_SL_BACKSTYLE:
		case DISPID_SL_BORDERSTYLE:
		case DISPID_SL_HALIGNMENT:
			c = 2;
			break;
		case DISPID_SL_APPEARANCE:
			c = 3;
			break;
		case DISPID_SL_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_SL_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<SysLink>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_SL_HALIGNMENT) && (iDescription == halCenter)) {
			propertyValue = halRight;
		}
		if((property == DISPID_SL_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP SysLink::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_SL_APPEARANCE:
		case DISPID_SL_BACKSTYLE:
		case DISPID_SL_BORDERSTYLE:
		case DISPID_SL_HALIGNMENT:
		case DISPID_SL_MOUSEPOINTER:
		case DISPID_SL_RIGHTTOLEFT:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<SysLink>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP SysLink::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_SL_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<SysLink>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT SysLink::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_SL_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SL_BACKSTYLE:
			switch(cookie) {
				case bstTransparent:
					description = GetResStringWithNumber(IDP_BACKSTYLETRANSPARENT, bstTransparent);
					break;
				case bstOpaque:
					description = GetResStringWithNumber(IDP_BACKSTYLEOPAQUE, bstOpaque);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SL_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SL_HALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HALIGNMENTLEFT, halLeft);
					break;
				/*case halCenter:
					description = GetResStringWithNumber(IDP_HALIGNMENTCENTER, halCenter);
					break;*/
				case halRight:
					description = GetResStringWithNumber(IDP_HALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SL_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_SL_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP SysLink::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 5;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_StockColorPage;
		pPropertyPages->pElems[3] = CLSID_StockFontPage;
		pPropertyPages->pElems[4] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP SysLink::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<SysLink>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP SysLink::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP SysLink::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = properties.hAcceleratorTable;
	pControlInfo->cAccel = static_cast<USHORT>(properties.hAcceleratorTable ? CopyAcceleratorTable(properties.hAcceleratorTable, NULL, 0) : 0);
	pControlInfo->dwFlags = 0;
	return S_OK;
}

STDMETHODIMP SysLink::OnMnemonic(LPMSG /*pMessage*/)
{
	if(GetStyle() & WS_DISABLED) {
		return S_OK;
	}

	SetFocus();
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT SysLink::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT SysLink::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_SL_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT SysLink::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP SysLink::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP SysLink::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_APPEARANCE);
	if(newValue < 0 || newValue > a3DLight) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_SL_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 17;
	return S_OK;
}

STDMETHODIMP SysLink::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"LabelControls");
	return S_OK;
}

STDMETHODIMP SysLink::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"LblCtls");
	return S_OK;
}

STDMETHODIMP SysLink::get_AutomaticallyMarkLinksAsVisited(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.automaticallyMarkLinksAsVisited);
	return S_OK;
}

STDMETHODIMP SysLink::put_AutomaticallyMarkLinksAsVisited(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_AUTOMATICALLYMARKLINKSASVISITED);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.automaticallyMarkLinksAsVisited != b) {
		properties.automaticallyMarkLinksAsVisited = b;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_SL_AUTOMATICALLYMARKLINKSASVISITED);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP SysLink::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_SL_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_SL_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_BackStyle(BackStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BackStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backStyle;
	return S_OK;
}

STDMETHODIMP SysLink::put_BackStyle(BackStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_BACKSTYLE);
	if(properties.backStyle != newValue) {
		properties.backStyle = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_SL_BACKSTYLE);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP SysLink::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_SL_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP SysLink::get_CaretLink(ILink** ppCaretLink)
{
	ATLASSERT_POINTER(ppCaretLink, ILink*);
	if(!ppCaretLink) {
		return E_POINTER;
	}

	if(IsWindow()) {
		LITEM link = {0};
		for(link.iLink = 0; ; link.iLink++) {
			link.mask = LIF_ITEMINDEX | LIF_STATE;
			link.stateMask = LIS_FOCUSED;
			if(SendMessage(LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
				if(link.state & LIS_FOCUSED) {
					ClassFactory::InitLink(link.iLink, this, IID_ILink, reinterpret_cast<LPUNKNOWN*>(ppCaretLink));
					break;
				}
			} else {
				break;
			}
		}
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP SysLink::putref_CaretLink(ILink* pNewCaretLink)
{
	PUTPROPPROLOG(DISPID_SL_CARETLINK);
	HRESULT hr = E_FAIL;

	LONG newCaretLink = -1;
	if(pNewCaretLink) {
		pNewCaretLink->get_Index(&newCaretLink);
		// TODO: Shouldn't we AddRef' pNewCaretLink?
	}

	if(IsWindow()) {
		hr = S_OK;
		LITEM link = {0};
		EnterSilentCaretChangeSection();
		for(link.iLink = 0; ; link.iLink++) {
			link.mask = LIF_ITEMINDEX | LIF_STATE;
			link.stateMask = LIS_FOCUSED;
			if(!SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
				break;
			}
		}
		link.iLink = newCaretLink;
		link.state = LIS_FOCUSED;
		SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
		Invalidate();
		LeaveSilentCaretChangeSection();

		if(newCaretLink != currentCaretLink) {
			Raise_CaretChanged(currentCaretLink, newCaretLink);
		}
	}

	FireOnChanged(DISPID_SL_CARETLINK);
	return hr;
}

STDMETHODIMP SysLink::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP SysLink::get_DetectDoubleClicks(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.detectDoubleClicks);
	return S_OK;
}

STDMETHODIMP SysLink::put_DetectDoubleClicks(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_DETECTDOUBLECLICKS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.detectDoubleClicks != b) {
		properties.detectDoubleClicks = b;
		SetDirty(TRUE);

		FireOnChanged(DISPID_SL_DETECTDOUBLECLICKS);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP SysLink::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
			if(IsWindow()) {
				if(static_cast<long>(newValue) & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
				}
				linkUnderMouse = -1;
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SL_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP SysLink::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_SL_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP SysLink::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_SL_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP SysLink::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_SL_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_FONT);
	return S_OK;
}

STDMETHODIMP SysLink::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_SL_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_FONT);
	return S_OK;
}

STDMETHODIMP SysLink::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP SysLink::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_SL_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_SL_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		switch(GetStyle() & (/*LWS_LEFT | LWS_CENTER | */LWS_RIGHT)) {
			/*case LWS_CENTER:
				properties.hAlignment = halCenter;
				break;*/
			case LWS_RIGHT:
				properties.hAlignment = halRight;
				break;
			case 0/*LWS_LEFT*/:
				properties.hAlignment = halLeft;
				break;
		}
	}

	*pValue = properties.hAlignment;
	return S_OK;
}

STDMETHODIMP SysLink::put_HAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_HALIGNMENT);
	if(properties.hAlignment != newValue) {
		properties.hAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_SL_HALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_HotTracking(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.hotTracking);
	return S_OK;
}

STDMETHODIMP SysLink::put_HotTracking(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_HOTTRACKING);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.hotTracking != b) {
		properties.hotTracking = b;
		SetDirty(TRUE);

		FireViewChange();
		FireOnChanged(DISPID_SL_HOTTRACKING);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP SysLink::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_SL_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SL_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP SysLink::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(toolTipStatus.hWndToolTip);
	}
	return S_OK;
}

STDMETHODIMP SysLink::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_SL_HWNDTOOLTIP);
	toolTipStatus.hWndToolTip = static_cast<HWND>(LongToHandle(newValue));

	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP SysLink::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP SysLink::get_Links(ILinks** ppLinks)
{
	ATLASSERT_POINTER(ppLinks, ILinks*);
	if(!ppLinks) {
		return E_POINTER;
	}

	ClassFactory::InitLinks(this, IID_ILinks, reinterpret_cast<LPUNKNOWN*>(ppLinks));
	return S_OK;
}

STDMETHODIMP SysLink::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP SysLink::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_SL_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP SysLink::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_SL_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP SysLink::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP SysLink::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SL_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP SysLink::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SL_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_ProcessReturnKey(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.processReturnKey = !(GetStyle() & LWS_IGNORERETURN);
	}

	*pValue = BOOL2VARIANTBOOL(properties.processReturnKey);
	return S_OK;
}

STDMETHODIMP SysLink::put_ProcessReturnKey(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_PROCESSRETURNKEY);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processReturnKey != b) {
		properties.processReturnKey = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.processReturnKey) {
				ModifyStyle(LWS_IGNORERETURN, 0);
			} else {
				ModifyStyle(0, LWS_IGNORERETURN);
			}
		}

		FireOnChanged(DISPID_SL_PROCESSRETURNKEY);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP SysLink::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP SysLink::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_SL_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP SysLink::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_SL_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
			}
			if(properties.rightToLeft & rtlText) {
				ModifyStyleEx(0, WS_EX_RTLREADING);
			} else {
				ModifyStyleEx(WS_EX_RTLREADING, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_SL_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_ShowToolTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(!toolTipStatus.hWndToolTip) {
			properties.showToolTips = FALSE;
		}
	}

	*pValue = BOOL2VARIANTBOOL(properties.showToolTips);
	return S_OK;
}

STDMETHODIMP SysLink::put_ShowToolTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_SHOWTOOLTIPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showToolTips != b) {
		properties.showToolTips = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showToolTips) {
				if(!toolTipStatus.hWndToolTip) {
					toolTipStatus.hWndToolTip = CreateWindowEx(0, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_NOPREFIX, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, *this, NULL, NULL, NULL);
					if(toolTipStatus.hWndToolTip) {
						TOOLINFO toolInfo = {0};
						toolInfo.cbSize = sizeof(TOOLINFO);
						toolInfo.uFlags = TTF_TRANSPARENT;
						toolInfo.hwnd = *this;
						toolInfo.lpszText = LPSTR_TEXTCALLBACK;
						GetClientRect(&toolInfo.rect);
						SendMessage(toolTipStatus.hWndToolTip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&toolInfo));
						SendMessage(toolTipStatus.hWndToolTip, WM_SETFONT, reinterpret_cast<WPARAM>(GetFont()), MAKELPARAM(TRUE, 0));
					}
				}
			} else if(::IsWindow(toolTipStatus.hWndToolTip)) {
				::DestroyWindow(toolTipStatus.hWndToolTip);
				toolTipStatus.hWndToolTip = NULL;
				toolTipStatus.toolTipLink = -1;
			}
		}
		FireOnChanged(DISPID_SL_SHOWTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP SysLink::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_SL_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP SysLink::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		GetWindowText(&properties.text);
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP SysLink::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_SL_TEXT);
	properties.text.AssignBSTR(newValue);
	if(IsWindow()) {
		SetWindowText(COLE2CT(properties.text));
	}
	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_TEXT);
	SendOnDataChange();
	return S_OK;
}

STDMETHODIMP SysLink::get_UseMnemonic(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.useMnemonic = !(GetStyle() & LWS_NOPREFIX);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useMnemonic);
	return S_OK;
}

STDMETHODIMP SysLink::put_UseMnemonic(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_USEMNEMONIC);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useMnemonic != b) {
		properties.useMnemonic = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.useMnemonic) {
				ModifyStyle(LWS_NOPREFIX, 0);
			} else {
				ModifyStyle(0, LWS_NOPREFIX);
			}
		}

		FireViewChange();
		FireOnChanged(DISPID_SL_USEMNEMONIC);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP SysLink::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_SL_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_UseVisualStyle(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.useVisualStyle = (GetStyle() & LWS_USEVISUALSTYLE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useVisualStyle);
	return S_OK;
}

STDMETHODIMP SysLink::put_UseVisualStyle(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_SL_USEVISUALSTYLE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useVisualStyle != b) {
		properties.useVisualStyle = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.useVisualStyle) {
				ModifyStyle(0, LWS_USEVISUALSTYLE);
			} else {
				ModifyStyle(LWS_USEVISUALSTYLE, 0);
			}
		}

		FireViewChange();
		FireOnChanged(DISPID_SL_USEVISUALSTYLE);
	}
	return S_OK;
}

STDMETHODIMP SysLink::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP SysLink::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP SysLink::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP SysLink::GetIdealSize(long maximumWidth/* = 0*/, OLE_XSIZE_PIXELS* pIdealWidth/* = NULL*/, OLE_YSIZE_PIXELS* pIdealHeight/* = NULL*/)
{
	if(IsWindow()) {
		if(maximumWidth == -1) {
			RECT clientRectangle = {0};
			GetClientRect(&clientRectangle);
			maximumWidth = clientRectangle.right - clientRectangle.left;
		}
		SIZE idealSize = {0};
		if(SendMessage(LM_GETIDEALSIZE, maximumWidth, reinterpret_cast<LPARAM>(&idealSize)) > 0) {
			RECT windowRectangle = {0, 0, idealSize.cx, idealSize.cy};
			AdjustWindowRectEx(&windowRectangle, GetStyle(), FALSE, GetExStyle());
			idealSize.cx = windowRectangle.right - windowRectangle.left;
			idealSize.cy = windowRectangle.bottom - windowRectangle.top;
			if(pIdealWidth) {
				*pIdealWidth = idealSize.cx;
			}
			if(pIdealHeight) {
				*pIdealHeight = idealSize.cy;
			}
			return S_OK;
		} else {
			return E_INVALIDARG;
		}
	}
	return E_FAIL;
}

STDMETHODIMP SysLink::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, ILink** ppHitLink)
{
	ATLASSERT_POINTER(ppHitLink, ILink*);
	if(!ppHitLink) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT flags = static_cast<UINT>(*pHitTestDetails);
		int linkIndex = HitTest(x, y, &flags);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(flags);
		}
		ClassFactory::InitLink(linkIndex, this, IID_ILink, reinterpret_cast<LPUNKNOWN*>(ppHitLink));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP SysLink::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP SysLink::Refresh(void)
{
	if(IsWindow()) {
		//dragDropStatus.HideDragImage();
		Invalidate();
		UpdateWindow();
		//dragDropStatus.ShowDragImage();
	}
	return S_OK;
}

STDMETHODIMP SysLink::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}


LRESULT SysLink::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT SysLink::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT SysLink::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(*this);
	}
	return lr;
}

LRESULT SysLink::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(*this);

	wasHandled = FALSE;
	return 0;
}

LRESULT SysLink::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT SysLink::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT SysLink::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<SysLink>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	if(-1 != currentCaretLink) {
		Raise_CaretChanged(currentCaretLink, -1);
	}
	return lr;
}

LRESULT SysLink::OnLButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}
	return 0;
}

LRESULT SysLink::OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	int newCaretLink = -1;
	CComPtr<ILink> pLink = NULL;
	if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
		LONG l = -1;
		pLink->get_Index(&l);
		newCaretLink = l;
	}
	if(newCaretLink != currentCaretLink) {
		Raise_CaretChanged(currentCaretLink, newCaretLink);
	}

	return lr;
}

LRESULT SysLink::OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT SysLink::OnMButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}
	return 0;
}

LRESULT SysLink::OnMButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	int newCaretLink = -1;
	CComPtr<ILink> pLink = NULL;
	if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
		LONG l = -1;
		pLink->get_Index(&l);
		newCaretLink = l;
	}
	if(newCaretLink != currentCaretLink) {
		Raise_CaretChanged(currentCaretLink, newCaretLink);
	}

	return lr;
}

LRESULT SysLink::OnMButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT SysLink::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT SysLink::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT SysLink::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);

	return 0;
}

LRESULT SysLink::OnMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
		UINT hitTestDetails = 0;
		int linkIndex = HitTest(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), &hitTestDetails);
		if(linkIndex != toolTipStatus.toolTipLink) {
			SendMessage(toolTipStatus.hWndToolTip, TTM_POP, 0, 0);
		}
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT SysLink::OnNCHitTest(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr == HTTRANSPARENT) {
		lr = HTCLIENT;
	}
	return lr;
}

LRESULT SysLink::OnNCMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}
	return 0;
}

LRESULT SysLink::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT SysLink::OnRButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}
	return 0;
}

LRESULT SysLink::OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	int newCaretLink = -1;
	CComPtr<ILink> pLink = NULL;
	if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
		LONG l = -1;
		pLink->get_Index(&l);
		newCaretLink = l;
	}
	if(newCaretLink != currentCaretLink) {
		Raise_CaretChanged(currentCaretLink, newCaretLink);
	}

	return lr;
}

LRESULT SysLink::OnRButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT SysLink::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		setCursor = (WindowFromPoint(mousePosition) == *this);
	}

	if(setCursor) {
		BOOL isOverLink = FALSE;
		LHITTESTINFO hitTestInfo = {0};
		hitTestInfo.pt = mousePosition;
		ScreenToClient(&hitTestInfo.pt);
		if(SendMessage(LM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo))) {
			isOverLink = TRUE;
		}

		if(!isOverLink && properties.mousePointer == mpDefault) {
			/* The SysLink control does not hit-test the position and always sets the hand cursor. It also
			 * handles WM_NCHITTEST in a way that makes anything that is not a link, transparent. For
			 * transparent parts we don't receive WM_SETCURSOR. But since we make those parets non-transparent in
			 * OnNCHitTest, the control receives WM_SETCURSOR messages for any point inside the client area, and
			 * therefore always sets the hand cursor. We hack around this behavior...
			 */
			// explicitly set the default cursor
			hCursor = static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
		} else if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT SysLink::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	int previousCaretLink = -1;
	CComPtr<ILink> pLink = NULL;
	if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
		LONG l = -1;
		pLink->get_Index(&l);
		previousCaretLink = l;
	}

	LRESULT lr = CComControl<SysLink>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<SysLink>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}

	if(m_bUIActive) {
		// if we have been entered via SHIFT+TAB, focus the last link
		if(GetKeyState(VK_SHIFT) & 0x8000) {
			if(GetKeyState(VK_TAB) & 0x8000) {
				int numberOfLinks = 0;
				int activeLink = -1;
				LITEM link = {0};
				for(link.iLink = 0; ; link.iLink++) {
					link.mask = LIF_ITEMINDEX | LIF_STATE;
					link.stateMask = LIS_FOCUSED;
					if(SendMessage(LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
						if(link.state & LIS_FOCUSED) {
							activeLink = link.iLink;
						}
					} else {
						numberOfLinks = link.iLink;
						break;
					}
				}
				EnterSilentCaretChangeSection();
				link.mask = LIF_ITEMINDEX | LIF_STATE;
				link.stateMask = LIS_FOCUSED;
				link.iLink = activeLink;
				link.state = 0;
				SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
				link.iLink = numberOfLinks - 1;
				link.state = LIS_FOCUSED;
				SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
				LeaveSilentCaretChangeSection();
			}
		} else {
			/* NOTE: The native SysLink would do this anyway, but then we would not detect that we need to raise
			 *       the CaretChanged event.
			 */
			if(GetKeyState(VK_TAB) & 0x8000) {
				int numberOfLinks = 0;
				int activeLink = -1;
				LITEM link = {0};
				for(link.iLink = 0; ; link.iLink++) {
					link.mask = LIF_ITEMINDEX | LIF_STATE;
					link.stateMask = LIS_FOCUSED;
					if(SendMessage(LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
						if(link.state & LIS_FOCUSED) {
							activeLink = link.iLink;
						}
					} else {
						numberOfLinks = link.iLink;
						break;
					}
				}
				EnterSilentCaretChangeSection();
				link.mask = LIF_ITEMINDEX | LIF_STATE;
				link.stateMask = LIS_FOCUSED;
				link.iLink = activeLink;
				link.state = 0;
				SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
				link.iLink = 0;
				link.state = LIS_FOCUSED;
				SendMessage(LM_SETITEM, 0, reinterpret_cast<LPARAM>(&link));
				LeaveSilentCaretChangeSection();
			}
		}
	}

	int newCaretLink = -1;
	pLink = NULL;
	if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
		LONG l = -1;
		pLink->get_Index(&l);
		newCaretLink = l;
	}
	if(newCaretLink != previousCaretLink) {
		Raise_CaretChanged(previousCaretLink, newCaretLink);
	}
	return lr;
}

LRESULT SysLink::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_SL_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(toolTipStatus.hWndToolTip) {
		SendMessage(toolTipStatus.hWndToolTip, WM_SETFONT, wParam, lParam);
	}

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_SL_FONT);
	}

	return lr;
}

LRESULT SysLink::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT SysLink::OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_SL_TEXT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		toolTipStatus.InvalidateToolTipLink();

		// update the accelerator table
		CComBSTR text;
		GetWindowText(&text);
		TCHAR newAccelerator = 0;
		for(long i = text.Length() - 1; i > 0; --i) {
			if((text[i - 1] == L'&') && (text[i] != L'&')) {
				COLE2T converter(text);
				newAccelerator = converter[i];
				break;
			}
		}

		if(newAccelerator != properties.accelerator) {
			properties.accelerator = newAccelerator;
			if(properties.hAcceleratorTable) {
				DestroyAcceleratorTable(properties.hAcceleratorTable);
			}

			// create a new accelerator table
			ACCEL accelerators[4] = {	{FALT, static_cast<WORD>(tolower(properties.accelerator)), 1},
																{0, static_cast<WORD>(tolower(properties.accelerator)), 1},
																{FVIRTKEY | FALT, LOBYTE(VkKeyScan(properties.accelerator)), 1},
																{FVIRTKEY | FALT | FSHIFT, LOBYTE(VkKeyScan(properties.accelerator)), 1} };
			properties.hAcceleratorTable = CreateAcceleratorTable(accelerators, _countof(accelerators));

			// report the new accelerator table to the container
			CComQIPtr<IOleControlSite, &IID_IOleControlSite> pSite(m_spClientSite);
			if(pSite) {
				pSite->OnControlInfoChanged();
			}
		}

		int newCaretLink = -1;
		CComPtr<ILink> pLink = NULL;
		if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
			LONG l = -1;
			pLink->get_Index(&l);
			newCaretLink = l;
		}
		if(newCaretLink != currentCaretLink) {
			Raise_CaretChanged(currentCaretLink, newCaretLink);
		}

		// analyzing what has changed would be too difficult, so always raise LinkMouseLeave and LinkMouseEnter
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		ScreenToClient(&mousePosition);
		UINT hitTestDetails = 0;
		int newLinkUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
		if(newLinkUnderMouse != linkUnderMouse) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			if(linkUnderMouse >= 0) {
				CComPtr<ILink> pLink = ClassFactory::InitLink(linkUnderMouse, this);
				Raise_LinkMouseLeave(pLink, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
			}

			linkUnderMouse = newLinkUnderMouse;
			if(linkUnderMouse >= 0) {
				CComPtr<ILink> pLink = ClassFactory::InitLink(linkUnderMouse, this);
				Raise_LinkMouseEnter(pLink, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
			}
		}

		if(!(properties.disabledEvents & deTextChangedEvents)) {
			Raise_TextChanged();
		}
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_SL_TEXT);
	SendOnDataChange();
	return lr;
}

LRESULT SysLink::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT SysLink::OnSize(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(toolTipStatus.hWndToolTip) {
		TOOLINFO toolInfo;
		if(properties.showToolTips) {
			toolTipStatus.InvalidateToolTipLink();
		}
		toolInfo.cbSize = sizeof(TOOLINFO);
		toolInfo.hwnd = *this;
		toolInfo.uId = 0;
		GetClientRect(&toolInfo.rect);
		SendMessage(toolTipStatus.hWndToolTip, TTM_NEWTOOLRECT, 0, reinterpret_cast<LPARAM>(&toolInfo));
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT SysLink::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT SysLink::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT SysLink::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	 *            not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	 *            even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		Invalidate();
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT SysLink::OnXButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}
	return 0;
}

LRESULT SysLink::OnXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	int newCaretLink = -1;
	CComPtr<ILink> pLink = NULL;
	if(SUCCEEDED(get_CaretLink(&pLink)) && pLink) {
		LONG l = -1;
		pLink->get_Index(&l);
		newCaretLink = l;
	}
	if(newCaretLink != currentCaretLink) {
		Raise_CaretChanged(currentCaretLink, newCaretLink);
	}

	return lr;
}

LRESULT SysLink::OnXButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(toolTipStatus.hWndToolTip) {
		toolTipStatus.RelayToToolTip(*this, message, wParam, lParam);
	}

	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT SysLink::OnReflectedCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.foreColor));
	SetBkMode(reinterpret_cast<HDC>(wParam), TRANSPARENT);
	if(properties.backStyle == bstOpaque) {
		if(!(properties.backColor & 0x80000000)) {
			SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
			SetDCBrushColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
			return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
		} else {
			SetBkColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.backColor));
			return reinterpret_cast<LRESULT>(GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
		}
	}

	RECT clientRectangle;
	::GetClientRect(reinterpret_cast<HWND>(lParam), &clientRectangle);
	FillRect(reinterpret_cast<HDC>(wParam), &clientRectangle, GetSysColorBrush(COLOR_3DFACE));
	if(RunTimeHelper::IsCommCtrl6() && flags.usingThemes) {
		BOOL useTransparentTextBackground = (DrawThemeParentBackground(reinterpret_cast<HWND>(lParam), reinterpret_cast<HDC>(wParam), &clientRectangle) == S_OK);
		// NOTE: DrawThemeParentBackground may reset the text color
		SetTextColor(reinterpret_cast<HDC>(wParam), OLECOLOR2COLORREF(properties.foreColor));
		SetBkMode(reinterpret_cast<HDC>(wParam), TRANSPARENT);
		if(!useTransparentTextBackground) {
			return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_3DFACE));
		}
	}
	return reinterpret_cast<LRESULT>(static_cast<HBRUSH>(GetStockObject(NULL_BRUSH)));
}

LRESULT SysLink::OnReflectedNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			if(reinterpret_cast<LPNMHDR>(lParam)->hwndFrom == *this) {
				return OnCustomDrawNotification(message, wParam, lParam, wasHandled);
			}
			break;
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT SysLink::OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	PLITEM pLinkData = reinterpret_cast<PLITEM>(lParam);
	ATLASSERT_POINTER(pLinkData, LITEM);
	if(!pLinkData) {
		wasHandled = FALSE;
		return TRUE;
	}

	if(pLinkData->mask & LIF_STATE) {
		if(pLinkData->stateMask & LIS_FOCUSED) {
			// caret link is being changed
			if(flags.silentCaretChanges == 0) {
				LRESULT lr = DefWindowProc(message, wParam, lParam);
				if(lr) {
					int newCaretLink;
					if(pLinkData->state & LIS_FOCUSED) {
						// caret has been set
						newCaretLink = pLinkData->iLink;
					} else {
						// caret has been cleared
						newCaretLink = -1;
					}
					if(newCaretLink != currentCaretLink) {
						Raise_CaretChanged(currentCaretLink, newCaretLink);
					}
				}
				return lr;
			}
		}
	}

	// let SysLink handle the message
	wasHandled = FALSE;
	return TRUE;
}


LRESULT SysLink::OnClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLINK* pDetails = reinterpret_cast<NMLINK*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLINK);
	if(!pDetails) {
		return 0;
	}

	CComPtr<ILink> pLink = ClassFactory::InitLink(pDetails->item.iLink, this);
	Raise_OpenLink(pLink, ecbClick);
	return 0;
}

LRESULT SysLink::OnCustomTextNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMCUSTOMTEXT pDetails = reinterpret_cast<LPNMCUSTOMTEXT>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMCUSTOMTEXT);
	if(!pDetails) {
		return 0;
	}

	CComBSTR text = pDetails->lpString;
	LONG maxTextLength = pDetails->nCount;

	VARIANT_BOOL drawText = VARIANT_TRUE;
	Raise_CustomizeTextDrawing(BOOL2VARIANTBOOL(pDetails->fLink), maxTextLength, &text, HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->lpRect), pDetails->uFormat, &drawText);

	int bufferSize = text.Length() + 1;
	if(bufferSize > pDetails->nCount + 1) {
		bufferSize = pDetails->nCount + 1;
	}
	if(bufferSize > 1) {
		lstrcpynW(const_cast<LPWSTR>(pDetails->lpString), text, bufferSize);
	} else {
		const_cast<LPWSTR>(pDetails->lpString)[0] = L'\0';
	}
	return !VARIANTBOOL2BOOL(drawText);
}

LRESULT SysLink::OnReturnNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	NMLINK* pDetails = reinterpret_cast<NMLINK*>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMLINK);
	if(!pDetails) {
		return 0;
	}
	// NOTE: LWS_IGNORERETURN does not seem to have any effect
	if(!(GetStyle() & LWS_IGNORERETURN)) {
		CComPtr<ILink> pLink = ClassFactory::InitLink(pDetails->item.iLink, this);
		Raise_OpenLink(pLink, ecbReturnKey);
	}
	return 0;
}

LRESULT SysLink::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	typedef BOOL WINAPI Str_SetPtrAFn(__deref_inout_opt LPSTR*, __in_opt LPCSTR);
	Str_SetPtrAFn* pfnStr_SetPtrA = NULL;
	HMODULE hComctl32DLL = LoadLibrary(TEXT("comctl32.dll"));
	if(hComctl32DLL) {
		pfnStr_SetPtrA = reinterpret_cast<Str_SetPtrAFn*>(GetProcAddress(hComctl32DLL, MAKEINTRESOURCEA(234)));
		FreeLibrary(hComctl32DLL);
	}

	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);

	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	VARIANT_BOOL cancelToolTip = VARIANT_FALSE;
	//if(linkIndex != toolTipStatus.toolTipLink) {
		if(pfnStr_SetPtrA) {
			pfnStr_SetPtrA(&toolTipStatus.pCurrentToolTipA, NULL);
		}
		toolTipStatus.toolTipLink = linkIndex;
		if(linkIndex >= 0) {
			CHAR pBuffer[INFOTIPSIZE];
			pBuffer[0] = 0;
			LPSTR pText = pBuffer;

			CComBSTR text = L"";
			CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);
			if(pLink) {
				pLink->get_URL(&text);
			}
			if(text.Length() == 0) {
				cancelToolTip = VARIANT_TRUE;
			}
			Raise_LinkGetInfoTipText(pLink, INFOTIPSIZE - 1, &text, &cancelToolTip);
			if(cancelToolTip == VARIANT_FALSE) {
				int bufferSize = text.Length() + 1;
				if(bufferSize > INFOTIPSIZE) {
					bufferSize = INFOTIPSIZE;
				}
				lstrcpynA(pBuffer, CW2A(text), bufferSize);
			}

			static const RECT margins = {4, 4, 4, 4};
			SendMessage(toolTipStatus.hWndToolTip, TTM_SETMARGIN, 0, reinterpret_cast<LPARAM>(&margins));
			HDC hDC = GetDC();
			int width = MulDiv(GetDeviceCaps(hDC, LOGPIXELSX), 300, 72);
			int maxWidth = GetDeviceCaps(hDC, HORZRES) * 3 / 4;
			ReleaseDC(hDC);
			SendMessage(toolTipStatus.hWndToolTip, TTM_SETMAXTIPWIDTH, 0, min(width, maxWidth));

			if(pfnStr_SetPtrA) {
				pfnStr_SetPtrA(&toolTipStatus.pCurrentToolTipA, pText);
			}
		}
	//}
	if(cancelToolTip == VARIANT_FALSE) {
		pDetails->lpszText = toolTipStatus.pCurrentToolTipA;
	}
	return 0;
}

LRESULT SysLink::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);

	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	VARIANT_BOOL cancelToolTip = VARIANT_FALSE;
	//if(linkIndex != toolTipStatus.toolTipLink) {
		Str_SetPtrW(&toolTipStatus.pCurrentToolTipW, NULL);
		toolTipStatus.toolTipLink = linkIndex;
		if(linkIndex >= 0) {
			WCHAR pBuffer[INFOTIPSIZE];
			pBuffer[0] = 0;
			LPWSTR pText = pBuffer;

			CComBSTR text = L"";
			CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);
			if(pLink) {
				pLink->get_URL(&text);
			}
			if(text.Length() == 0) {
				cancelToolTip = VARIANT_TRUE;
			}
			Raise_LinkGetInfoTipText(pLink, INFOTIPSIZE - 1, &text, &cancelToolTip);
			if(cancelToolTip == VARIANT_FALSE) {
				int bufferSize = text.Length() + 1;
				if(bufferSize > INFOTIPSIZE) {
					bufferSize = INFOTIPSIZE;
				}
				lstrcpynW(pBuffer, text, bufferSize);
			}

			static const RECT margins = {4, 4, 4, 4};
			SendMessage(toolTipStatus.hWndToolTip, TTM_SETMARGIN, 0, reinterpret_cast<LPARAM>(&margins));
			HDC hDC = GetDC();
			int width = MulDiv(GetDeviceCaps(hDC, LOGPIXELSX), 300, 72);
			int maxWidth = GetDeviceCaps(hDC, HORZRES) * 3 / 4;
			ReleaseDC(hDC);
			SendMessage(toolTipStatus.hWndToolTip, TTM_SETMAXTIPWIDTH, 0, min(width, maxWidth));

			Str_SetPtrW(&toolTipStatus.pCurrentToolTipW, pText);
		}
	//}
	if(cancelToolTip == VARIANT_FALSE) {
		pDetails->lpszText = toolTipStatus.pCurrentToolTipW;
	}

	return 0;
}


LRESULT SysLink::OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMCUSTOMDRAW pDetails = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
	ATLASSERT_POINTER(pDetails, NMCUSTOMDRAW);
	if(!pDetails) {
		return DefWindowProc(message, wParam, lParam);
	}
	if(!(pDetails->dwDrawStage & CDDS_ITEM)) {
		customDrawNonLinkPartIndex = -1;
	}

	ATLASSERT((pDetails->uItemState & CDIS_NEARHOT) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_OTHERSIDEHOT) == 0);
	ATLASSERT((pDetails->uItemState & CDIS_DROPHILITED) == 0);

	CustomDrawReturnValuesConstants returnValue = static_cast<CustomDrawReturnValuesConstants>(this->DefWindowProc(message, wParam, lParam));
	
	if(properties.automaticallyMarkLinksAsVisited) {
		if(pDetails->dwDrawStage == CDDS_PREPAINT) {
			returnValue = cdrvNotifyLinkDraw;
		} else if(pDetails->dwDrawStage == CDDS_ITEMPREPAINT) {
			if(pDetails->dwItemSpec >= 0) {
				LITEM link = {0};
				link.iLink = pDetails->dwItemSpec;
				link.mask = LIF_ITEMINDEX | LIF_STATE;
				link.stateMask = LIS_VISITED | LIS_ENABLED | LIS_DEFAULTCOLORS;
				if(SendMessage(LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
					if((link.state & (LIS_VISITED | LIS_ENABLED | LIS_DEFAULTCOLORS)) == (LIS_VISITED | LIS_ENABLED)) {
						CRegKeyEx key;
						if(key.Open(HKEY_CURRENT_USER, TEXT("Software\\Microsoft\\Internet Explorer\\Settings")) == ERROR_SUCCESS) {
							TCHAR pValue[12] = {0};
							ULONG count = 12;
							COLORREF color = CLR_INVALID;
							BOOL useColor = TRUE;
							if(key.QueryStringValue(TEXT("Anchor Color Visited"), pValue, &count) == ERROR_SUCCESS) {
								LPTSTR pColorStr = pValue;
								int c[3] = {-1, -1, -1};
								LPTSTR p = NULL;
								for(int i = 0; i < 2; i++) {
									for(p = pColorStr; *p != TEXT('\0'); p = CharNext(p)) {
										if(*p == TEXT(',')) {
											*p = TEXT('\0');
											ConvertStringToInteger(pColorStr, c[i]);
											pColorStr = &p[1];
											break;
										}
									}
									if(c[i] == -1) {
										useColor = FALSE;
										break;
									}
								}
								if(*pColorStr == TEXT('\0')) {
									useColor = FALSE;
								}
								ConvertStringToInteger(pColorStr, c[2]);
								if(useColor) {
									color = RGB(c[0], c[1], c[2]);
									ATLASSERT(color != CLR_INVALID);
									if(color != CLR_INVALID) {
										SetTextColor(pDetails->hdc, color);
									}
								}
							}
						}
					}
				}
			}
		}
	}
	
	if(!(properties.disabledEvents & deCustomDraw)) {
		CComPtr<ILink> pLink = NULL;
		OLE_COLOR colorText = 0;
		OLE_COLOR colorTextBackground = 0;
		if(pDetails->dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
			pLink = ClassFactory::InitLink(pDetails->dwItemSpec, this);
			colorText = GetTextColor(pDetails->hdc);
			colorTextBackground = GetBkColor(pDetails->hdc);
		}
		if(!pLink && (pDetails->dwDrawStage & CDDS_ITEM)) {
			if(pDetails->dwDrawStage == CDDS_ITEMPREPAINT || pDetails->dwDrawStage == CDDS_ITEMPREERASE) {
				++customDrawNonLinkPartIndex;
			}
		}
		Raise_CustomDraw(pLink, (pLink ? -1 : customDrawNonLinkPartIndex), &colorText, &colorTextBackground, static_cast<CustomDrawStageConstants>(pDetails->dwDrawStage), static_cast<CustomDrawLinkStateConstants>(pDetails->uItemState), HandleToLong(pDetails->hdc), reinterpret_cast<RECTANGLE*>(&pDetails->rc), &returnValue);

		if(pDetails->dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
			SetTextColor(pDetails->hdc, colorText);
			SetBkColor(pDetails->hdc, colorTextBackground);
		}
	}

	return static_cast<LRESULT>(returnValue);
}


void SysLink::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
			properties.font.owningFontResource = FALSE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT SysLink::Raise_CaretChanged(int previousCaretLink, int newCaretLink)
{
	if(newCaretLink == previousCaretLink) {
		return S_OK;
	}

	currentCaretLink = newCaretLink;
	if(flags.silentCaretChanges == 0) {
		//if(m_nFreezeEvents == 0) {
			CComPtr<ILink> pSLPreviousCaretLink = ClassFactory::InitLink(previousCaretLink, this);
			CComPtr<ILink> pSLNewCaretLink = ClassFactory::InitLink(newCaretLink, this);
			return Fire_CaretChanged(pSLPreviousCaretLink, pSLNewCaretLink);
		//}
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedLink = HitTest(x, y, &hitTestDetails);
	CComPtr<ILink> pLink = ClassFactory::InitLink(mouseStatus.lastClickedLink, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_Click(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		BOOL dontUsePosition = FALSE;
		LONG linkIndex = -1;
		UINT hitTestDetails = 0;
		if(x == -1 && y == -1) {
			// the event was caused by the keyboard
			dontUsePosition = TRUE;
			if(properties.processContextMenuKeys) {
				// retrieve the caret link and propose its rectangle's middle as the menu's position
				WTL::CRect clientRectangle;
				GetClientRect(&clientRectangle);
				WTL::CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
				hitTestDetails = htTextOrBackground;

				CComPtr<ILink> pCaretLink = NULL;
				if(SUCCEEDED(get_CaretLink(&pCaretLink)) && pCaretLink) {
					pCaretLink->get_Index(&linkIndex);
					hitTestDetails = htLink;

					// unfortunately it is not possible to retrieve the link's rectangle
				}
			} else {
				return S_OK;
			}
		}

		if(!dontUsePosition) {
			linkIndex = HitTest(x, y, &hitTestDetails);
		}
		CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

		return Fire_ContextMenu(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_CustomDraw(ILink* pLink, LONG nonLinkPartIndex, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawLinkStateConstants linkState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomDraw(pLink, nonLinkPartIndex, pTextColor, pTextBackColor, drawStage, linkState, hDC, pDrawingRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_CustomizeTextDrawing(VARIANT_BOOL isLink, LONG maxTextLength, BSTR* pText, LONG hDC, RECTANGLE* pDrawingRectangle, LONG drawTextFlags, VARIANT_BOOL* pDrawText)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomizeTextDrawing(isLink, maxTextLength, pText, hDC, pDrawingRectangle, drawTextFlags, pDrawText);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(x, y, &hitTestDetails);
	if(linkIndex != mouseStatus.lastClickedLink) {
		linkIndex = -1;
	}
	mouseStatus.lastClickedLink = -1;
	CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_DestroyedControlWindow(HWND hWnd)
{
	KillTimer(timers.ID_REDRAW);

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_LinkGetInfoTipText(ILink* pLink, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_LinkGetInfoTipText(pLink, maxInfoTipLength, pInfoTipText, pAbortToolTip);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_LinkMouseEnter(ILink* pLink, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(pLink && properties.hotTracking) {
		pLink->put_Hot(VARIANT_TRUE);
	}
	if(!(properties.disabledEvents & deMouseEvents)) {
		if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
			return Fire_LinkMouseEnter(pLink, button, shift, x, y, hitTestDetails);
		}
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_LinkMouseLeave(ILink* pLink, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
{
	if(pLink && properties.hotTracking) {
		pLink->put_Hot(VARIANT_FALSE);
	}
	if(!(properties.disabledEvents & deMouseEvents)) {
		if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
			return Fire_LinkMouseLeave(pLink, button, shift, x, y, hitTestDetails);
		}
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedLink = HitTest(x, y, &hitTestDetails);
	CComPtr<ILink> pLink = ClassFactory::InitLink(mouseStatus.lastClickedLink, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(x, y, &hitTestDetails);
	if(linkIndex != mouseStatus.lastClickedLink) {
		linkIndex = -1;
	}
	mouseStatus.lastClickedLink = -1;
	CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL fireMouseDown = FALSE;
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		fireMouseDown = TRUE;
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
	}
	mouseStatus.StoreClickCandidate(button);
	SetCapture();

	if(mouseStatus.IsDoubleClickCandidate(button)) {
		/* Watch for double-clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the double-click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_DblClick(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RDblClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MDblClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XDblClick(button, shift, x, y);
					}
					break;
			}
			mouseStatus.RemoveAllDoubleClickCandidates();
		}

		mouseStatus.RemoveClickCandidate(button);
		if(GetCapture() == *this) {
			ReleaseCapture();
		}
		fireMouseDown = FALSE;
	} else {
		mouseStatus.RemoveAllDoubleClickCandidates();
	}

	HRESULT hr = S_OK;
	if(fireMouseDown) {
		UINT hitTestDetails = 0;
		int linkIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

		hr = Fire_MouseDown(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return hr;
}

inline HRESULT SysLink::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		UINT hitTestDetails = 0;
		int linkIndex = HitTest(x, y, &hitTestDetails);
		linkUnderMouse = linkIndex;
		CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

		return Fire_MouseEnter(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		UINT hitTestDetails = 0;
		int linkIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

		return Fire_MouseHover(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<ILink> pLink = ClassFactory::InitLink(linkUnderMouse, this);
	if(pLink) {
		Raise_LinkMouseLeave(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	linkUnderMouse = -1;

	mouseStatus.LeaveControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		pLink = ClassFactory::InitLink(linkIndex, this);
		return Fire_MouseLeave(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int linkIndex = HitTest(x, y, &hitTestDetails);
	CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

	// Do we move over another item than before?
	if(linkIndex != linkUnderMouse) {
		CComPtr<ILink> pPrevLink = ClassFactory::InitLink(linkUnderMouse, this);
		if(pPrevLink) {
			Raise_LinkMouseLeave(pPrevLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
		linkUnderMouse = linkIndex;
		if(pLink) {
			Raise_LinkMouseEnter(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		}
	}
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseMove(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT SysLink::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		UINT hitTestDetails = 0;
		int linkIndex = HitTest(x, y, &hitTestDetails);
		CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

		hr = Fire_MouseUp(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
			if(properties.detectDoubleClicks) {
				mouseStatus.StoreDoubleClickCandidate(button);
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT SysLink::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	CComPtr<ILink> pDropTarget = ClassFactory::InitLink(dropTarget, this);

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
				always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();

	return hr;
}

inline HRESULT SysLink::Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	CComPtr<ILink> pDropTarget = ClassFactory::InitLink(dropTarget, this);

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	return hr;
}

inline HRESULT SysLink::Raise_OLEDragLeave()
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
	CComPtr<ILink> pDropTarget = ClassFactory::InitLink(dropTarget, this);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();

	return hr;
}

inline HRESULT SysLink::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	CComPtr<ILink> pDropTarget = ClassFactory::InitLink(dropTarget, this);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	return hr;
}

inline HRESULT SysLink::Raise_OpenLink(ILink* pLink, EventCausedByConstants causedBy)
{
	if(pLink) {
		pLink->put_Visited(VARIANT_TRUE);
	}
	//if(m_nFreezeEvents == 0) {
		return Fire_OpenLink(pLink, causedBy);
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedLink = HitTest(x, y, &hitTestDetails);
	CComPtr<ILink> pLink = ClassFactory::InitLink(mouseStatus.lastClickedLink, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(x, y, &hitTestDetails);
	if(linkIndex != mouseStatus.lastClickedLink) {
		linkIndex = -1;
	}
	mouseStatus.lastClickedLink = -1;
	CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_RecreatedControlWindow(HWND hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(HandleToLong(hWnd));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_TextChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TextChanged();
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedLink = HitTest(x, y, &hitTestDetails);
	CComPtr<ILink> pLink = ClassFactory::InitLink(mouseStatus.lastClickedLink, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT SysLink::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int linkIndex = HitTest(x, y, &hitTestDetails);
	if(linkIndex != mouseStatus.lastClickedLink) {
		linkIndex = -1;
	}
	mouseStatus.lastClickedLink = -1;
	CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, this);

	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(pLink, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void SysLink::EnterSilentCaretChangeSection(void)
{
	++flags.silentCaretChanges;
}

void SysLink::LeaveSilentCaretChangeSection(void)
{
	--flags.silentCaretChanges;
}


void SysLink::RecreateControlWindow(void)
{
	if(m_bInPlaceActive && !flags.dontRecreate) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

DWORD SysLink::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD SysLink::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(!properties.processReturnKey) {
		style |= LWS_IGNORERETURN;
	}
	if(IsComctl32Version610OrNewer()) {
		switch(properties.hAlignment) {
			case halLeft:
				style |= 0/*LWS_LEFT*/;
				break;
			/*case halCenter:
				style |= LWS_CENTER;
				break;*/
			case halRight:
				style |= LWS_RIGHT;
				break;
		}
		if(!properties.useMnemonic) {
			style |= LWS_NOPREFIX;
		}
		if(!(properties.disabledEvents & deCustomizeTextDrawing)) {
			style |= LWS_USECUSTOMTEXT;
		}
		if(properties.useVisualStyle) {
			style |= LWS_USEVISUALSTYLE;
		}
	}
	return style;
}

void SysLink::SendConfigurationMessages(void)
{
	SetWindowText(COLE2CT(properties.text));

	ApplyFont();

	if(properties.showToolTips) {
		toolTipStatus.hWndToolTip = CreateWindowEx(0, TOOLTIPS_CLASS, NULL, WS_POPUP | TTS_NOPREFIX, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, *this, NULL, NULL, NULL);
		if(toolTipStatus.hWndToolTip) {
			TOOLINFO toolInfo = {0};
			toolInfo.cbSize = sizeof(TOOLINFO);
			toolInfo.uFlags = TTF_TRANSPARENT;
			toolInfo.hwnd = *this;
			toolInfo.lpszText = LPSTR_TEXTCALLBACK;
			GetClientRect(&toolInfo.rect);
			SendMessage(toolTipStatus.hWndToolTip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&toolInfo));
			SendMessage(toolTipStatus.hWndToolTip, WM_SETFONT, reinterpret_cast<WPARAM>(GetFont()), MAKELPARAM(TRUE, 0));
		}
	}
}

HCURSOR SysLink::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

int SysLink::HitTest(LONG x, LONG y, UINT* pFlags)
{
	ATLASSERT(IsWindow());

	UINT flags = 0;
	if(pFlags) {
		flags = *pFlags;
	}
	LHITTESTINFO hitTestInfo = {{x, y}, {0}};
	int linkIndex = -1;

	// Are we outside the window?
	WTL::CRect windowRectangle;
	GetWindowRect(&windowRectangle);
	ScreenToClient(&windowRectangle);
	if(windowRectangle.PtInRect(hitTestInfo.pt)) {
		hitTestInfo.item.iLink = -1;
		if(SendMessage(LM_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo))) {
			linkIndex = hitTestInfo.item.iLink;
			if(linkIndex >= 0) {
				flags |= SLHT_LINK;
			} else {
				flags |= SLHT_TEXTORBACKGROUND;
			}
		} else {
			flags |= SLHT_TEXTORBACKGROUND;
		}
	} else {
		if(x < windowRectangle.left) {
			flags |= SLHT_TOLEFT;
		} else if(x > windowRectangle.right) {
			flags |= SLHT_TORIGHT;
		}
		if(y < windowRectangle.top) {
			flags |= SLHT_ABOVE;
		} else if(y > windowRectangle.bottom) {
			flags |= SLHT_BELOW;
		}
	}

	if(pFlags) {
		*pFlags = flags;
	}
	return linkIndex;
}

BOOL SysLink::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}


BOOL SysLink::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}