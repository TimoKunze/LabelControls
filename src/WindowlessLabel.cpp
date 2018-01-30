// WindowlessLabel.cpp: Superclasses Static.

#include "stdafx.h"
#include "WindowlessLabel.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC WindowlessLabel::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};

LONG WindowlessLabel::timerMapReferences = 0;
CRITICAL_SECTION WindowlessLabel::timerMapCriticalSection;
#ifdef USE_STL
	std::map<UINT_PTR, WindowlessLabel*> WindowlessLabel::timerMap;
#else
	CAtlMap<UINT_PTR, WindowlessLabel*> WindowlessLabel::timerMap;
#endif


WindowlessLabel::WindowlessLabel()
{
	properties.font.InitializePropertyWatcher(this, DISPID_WLLBL_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_WLLBL_MOUSEICON);

	SIZEL size = {100, 20};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	// we prefer being windowless
	m_bWindowOnly = FALSE;

	// initialize
	mouseHoverTimerID = 0;
	mouseLeaveTimerID = 0;
	mightHaveLeftControl = FALSE;
	if(InterlockedIncrement(&timerMapReferences) == 1) {
		InitializeCriticalSection(&timerMapCriticalSection);
	}
}

WindowlessLabel::~WindowlessLabel()
{
	if(mouseHoverTimerID) {
		::KillTimer(NULL, mouseHoverTimerID);
		mouseHoverTimerID = 0;
	}
	if(mouseLeaveTimerID) {
		::KillTimer(NULL, mouseLeaveTimerID);
		mouseLeaveTimerID = 0;
	}
	if(InterlockedDecrement(&timerMapReferences) == 0) {
		DeleteCriticalSection(&timerMapCriticalSection);
	}
}

DWORD WindowlessLabel::_GetViewStatus()
{
	switch(properties.backStyle) {
		case bstTransparent:
			return VIEWSTATUS_OPAQUE | VIEWSTATUS_DVASPECTOPAQUE | VIEWSTATUS_DVASPECTTRANSPARENT;					// required to make the control work correctly inside VB6 PictureBox with AutoRedraw=True
			break;
		case bstOpaque:
			return VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE;
			break;
	}
	return VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE;
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP WindowlessLabel::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IWindowlessLabel, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP WindowlessLabel::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<WindowlessLabel>::Load(pPropertyBag, pErrorLog);
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

STDMETHODIMP WindowlessLabel::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<WindowlessLabel>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
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

STDMETHODIMP WindowlessLabel::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 14 VT_I4 properties...
	pSize->QuadPart += 14 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 7 VT_BOOL properties...
	pSize->QuadPart += 7 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
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

STDMETHODIMP WindowlessLabel::Load(LPSTREAM pStream)
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
	if(subSignature != 0x03030303/*4x 0x03 (-> WindowlessLabel)*/) {
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AutoSizeConstants autoSize = static_cast<AutoSizeConstants>(propertyValue.lVal);
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
	VARIANT_BOOL clipLastLine = propertyValue.boolVal;
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
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	TextTruncationStyleConstants textTruncationStyle = static_cast<TextTruncationStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useMnemonic = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	VAlignmentConstants vAlignment = static_cast<VAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	WordWrappingConstants wordWrapping = static_cast<WordWrappingConstants>(propertyValue.lVal);

	BOOL b = flags.dontRecreate;
	flags.dontRecreate = TRUE;
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoSize(autoSize);
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
	hr = put_ClipLastLine(clipLastLine);
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
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
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
	hr = put_TextTruncationStyle(textTruncationStyle);
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
	hr = put_VAlignment(vAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_WordWrapping(wordWrapping);
	if(FAILED(hr)) {
		return hr;
	}
	flags.dontRecreate = b;
	RecreateControlWindow();

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG subSignature = 0x03030303/*4x 0x03 (-> WindowlessLabel)*/;
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
	propertyValue.lVal = properties.autoSize;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.clipLastLine);
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

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
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
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.textTruncationStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
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
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.vAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.wordWrapping;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HRESULT WindowlessLabel::OnDraw(ATL_DRAWINFO& drawInfo)
{
	CDCHandle targetDC = drawInfo.hdcDraw;
	RECT drawingRectangle = *reinterpret_cast<LPCRECT>(drawInfo.prcBounds);
	Draw(targetDC, drawingRectangle);
	return S_OK;
}

void WindowlessLabel::Draw(CDCHandle& targetDC, RECT& drawingRectangle)
{
	if(properties.font.currentFont.IsNull()) {
		ApplyFont();
	}

	DWORD previousLayout = GetLayout(targetDC);
	BOOL restoreLayout = FALSE;
	if(previousLayout & LAYOUT_RTL) {
		int tmp = drawingRectangle.left;
		drawingRectangle.left = drawingRectangle.right + 1;
		drawingRectangle.right = tmp + 1;
		SetLayout(targetDC, 0);
		restoreLayout = TRUE;
	}
	int dcState = targetDC.SaveDC();
	targetDC.SelectFont(properties.font.currentFont);
	targetDC.SetBkMode(TRANSPARENT);
	targetDC.SetTextColor(OLECOLOR2COLORREF(properties.foreColor));

	if(!properties.dontRedraw || IsInDesignMode()) {
		RECT contentRectangle = drawingRectangle;
		DrawBorder(targetDC, drawingRectangle, contentRectangle);
		DrawBackground(targetDC, contentRectangle);

		HFONT hFontWithoutAntialiasing = NULL;
		HFONT hNormalFont = targetDC.GetCurrentFont();
		if(!properties.enabled) {
			LOGFONT font = {0};
			if(GetObject(hNormalFont, sizeof(font), &font)) {
				DWORD quality = DEFAULT_QUALITY | NONANTIALIASED_QUALITY;
				hFontWithoutAntialiasing = CreateFont(font.lfHeight, 0, 0, 0, font.lfWeight, font.lfItalic, font.lfUnderline, font.lfStrikeOut, font.lfCharSet, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, quality, FF_DONTCARE, font.lfFaceName);
			}
			if(hFontWithoutAntialiasing) {
				targetDC.SelectFont(hFontWithoutAntialiasing);
			}

			targetDC.SetTextColor(GetSysColor(COLOR_3DHIGHLIGHT));
			RECT rc = contentRectangle;
			OffsetRect(&rc, (properties.rightToLeft & rtlLayout ? -1 : 1), 1);
			DrawCaption(targetDC, rc);
			targetDC.SetTextColor(GetSysColor(COLOR_3DSHADOW));
		}
		DrawCaption(targetDC, contentRectangle);
		if(hFontWithoutAntialiasing) {
			targetDC.SelectFont(hNormalFont);
			DeleteObject(hFontWithoutAntialiasing);
		}
	}

	targetDC.RestoreDC(dcState);
	if(restoreLayout) {
		SetLayout(targetDC, previousLayout);
	}
}

void WindowlessLabel::DrawBorder(CDCHandle& targetDC, RECT& drawingRectangle, RECT& contentRectangle)
{
	DWORD style = (properties.borderStyle == bsFixedSingle ? WS_BORDER : 0);
	DWORD extendedStyle = 0;
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

	contentRectangle = drawingRectangle;
	if(style & WS_BORDER) {
		int width = GetSystemMetrics(SM_CXBORDER);
		int height = GetSystemMetrics(SM_CYBORDER);
		targetDC.SelectBrush(GetSysColorBrush(extendedStyle & WS_EX_CLIENTEDGE ? COLOR_3DFACE : COLOR_WINDOWFRAME));
		targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, height, PATCOPY);
		targetDC.PatBlt(contentRectangle.left, contentRectangle.top, width, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom, contentRectangle.right - contentRectangle.left, -height, PATCOPY);
		targetDC.PatBlt(contentRectangle.right, contentRectangle.top, -width, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		InflateRect(&contentRectangle, -width, -height);
	}
	if(extendedStyle & WS_EX_CLIENTEDGE) {
		targetDC.SelectBrush(GetSysColorBrush(COLOR_BTNSHADOW));
		if(extendedStyle & WS_EX_LAYOUTRTL) {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.right - 1, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		} else {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		}
		targetDC.SelectBrush(GetSysColorBrush(COLOR_BTNHIGHLIGHT));
		if(extendedStyle & WS_EX_LAYOUTRTL) {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom - 1, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		} else {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom - 1, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.right - 1, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		}
		InflateRect(&contentRectangle, -1, -1);
		targetDC.SelectBrush(GetSysColorBrush(COLOR_3DDKSHADOW));
		if(extendedStyle & WS_EX_LAYOUTRTL) {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.right - 1, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		} else {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		}
		targetDC.SelectBrush(GetSysColorBrush(COLOR_3DLIGHT));
		if(extendedStyle & WS_EX_LAYOUTRTL) {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom - 1, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		} else {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom - 1, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.right - 1, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		}
		InflateRect(&contentRectangle, -1, -1);
	}
	if(extendedStyle & WS_EX_STATICEDGE) {
		targetDC.SelectBrush(GetSysColorBrush(COLOR_BTNSHADOW));
		if(extendedStyle & WS_EX_LAYOUTRTL) {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.right - 1, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		} else {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		}
		targetDC.SelectBrush(GetSysColorBrush(COLOR_BTNHIGHLIGHT));
		if(extendedStyle & WS_EX_LAYOUTRTL) {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom - 1, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.left, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		} else {
			targetDC.PatBlt(contentRectangle.left, contentRectangle.bottom - 1, contentRectangle.right - contentRectangle.left, 1, PATCOPY);
			targetDC.PatBlt(contentRectangle.right - 1, contentRectangle.top, 1, contentRectangle.bottom - contentRectangle.top, PATCOPY);
		}
		InflateRect(&contentRectangle, -1, -1);
	}
}

void WindowlessLabel::DrawBackground(CDCHandle& targetDC, RECT& contentRectangle)
{
	if(properties.backStyle == bstOpaque) {
		if(!(properties.backColor & 0x80000000)) {
			HBRUSH hDCBrush = static_cast<HBRUSH>(GetStockObject(DC_BRUSH));
			HBRUSH hPreviousBrush = targetDC.SelectBrush(hDCBrush);
			COLORREF previousColor = targetDC.SetDCBrushColor(properties.backColor);
			targetDC.FillRect(&contentRectangle, hDCBrush);
			targetDC.SetDCBrushColor(previousColor);
			targetDC.SelectBrush(hPreviousBrush);
		} else {
			targetDC.FillRect(&contentRectangle, GetSysColorBrush(properties.backColor & 0x0FFFFFFF));
		}
	}
}

void WindowlessLabel::DrawCaption(CDCHandle& targetDC, RECT& contentRectangle)
{
	int textLength = properties.text.Length();
	if(textLength > 0) {
		DWORD extendedStyle = 0;
		if(properties.rightToLeft & rtlLayout) {
			extendedStyle |= WS_EX_LAYOUTRTL;
		}

		DWORD drawTextFlags = 0;
		if(!properties.clipLastLine) {
			drawTextFlags |= DT_EDITCONTROL;
		}
		switch(properties.hAlignment) {
			case halLeft:
				drawTextFlags |= (extendedStyle & WS_EX_LAYOUTRTL ? DT_RIGHT : DT_LEFT);
				break;
			case halCenter:
				drawTextFlags |= DT_CENTER;
				break;
			case halRight:
				drawTextFlags |= (extendedStyle & WS_EX_LAYOUTRTL ? DT_LEFT : DT_RIGHT);
				break;
		}
		if(properties.rightToLeft & rtlText) {
			drawTextFlags |= DT_RTLREADING;
		}
		switch(properties.textTruncationStyle) {
			case ttsEndEllipsis:
				drawTextFlags |= DT_END_ELLIPSIS;
				break;
			case ttsWordEllipsis:
				drawTextFlags |= DT_WORD_ELLIPSIS;
				break;
			case ttsPathEllipsis:
				drawTextFlags |= DT_PATH_ELLIPSIS;
				break;
		}
		if(!properties.useMnemonic) {
			drawTextFlags |= DT_NOPREFIX;
		}
		switch(properties.wordWrapping) {
			case wwAutomatic:
				drawTextFlags |= DT_WORDBREAK;
				break;
			case wwSingleLine:
				drawTextFlags |= DT_SINGLELINE;
				break;
		}
		if(drawTextFlags & DT_RIGHT) {
			contentRectangle.right--;
		}

		if(properties.vAlignment != valTop) {
			if(drawTextFlags & DT_SINGLELINE) {
				switch(properties.vAlignment) {
					case valTop:
						drawTextFlags |= DT_TOP;
						break;
					case valCenter:
						drawTextFlags |= DT_VCENTER;
						break;
					case valBottom:
						drawTextFlags |= DT_BOTTOM;
						break;
				}
			} else {
				COLE2T converter(properties.text);
				CRect requiredRectangle = contentRectangle;
				requiredRectangle.bottom = 0x7FFFFFFF;
				targetDC.DrawText(converter, textLength, &requiredRectangle, drawTextFlags | DT_TOP | DT_CALCRECT);

				switch(properties.vAlignment) {
					case valCenter:
						contentRectangle.top -= (requiredRectangle.Height() - (contentRectangle.bottom - contentRectangle.top)) / 2;			// don't bit-shift here because the difference can be negative
						contentRectangle.bottom = contentRectangle.top + requiredRectangle.Height();
						break;
					case valBottom:
						contentRectangle.top = contentRectangle.bottom - requiredRectangle.Height();
						break;
				}
				drawTextFlags |= DT_TOP;
			}
		}

		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		BOOL drawAccel = FALSE;
		if(hWndParent) {
			drawAccel = ((SendMessage(hWndParent, WM_QUERYUISTATE, 0, 0) & UISF_HIDEACCEL) == 0);
		}
		if(!drawAccel) {
			drawTextFlags |= DT_HIDEPREFIX;
		}
		COLE2T converter(properties.text);
		targetDC.DrawText(converter, textLength, &contentRectangle, drawTextFlags);
	}
}

void WindowlessLabel::Redraw(void)
{
	if(!properties.dontRedraw || IsInDesignMode()) {
		BOOL didRefresh = FALSE;
		if(properties.backStyle == bstTransparent) {
			if(!didRefresh && m_spInPlaceSite) {
				/* NOTE: Use InvalidateRgn instead of InvalidateRect!! InvalidateRect won't erase the background
				 *       even if you tell it to do it! */
				didRefresh = SUCCEEDED(m_spInPlaceSite->InvalidateRgn(NULL, TRUE));
			}
		}
		if(!didRefresh) {
			FireViewChange();
		}
	}
}

STDMETHODIMP WindowlessLabel::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<WindowlessLabel>::IOleObject_SetClientSite(pClientSite);

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
		} else {
			SysFreeString(buffer);
		}
	}

	return hr;
}

STDMETHODIMP WindowlessLabel::IOleInPlaceObject_InPlaceDeactivate(void)
{
	if(m_spInPlaceSite && flags.hasSubclassedParentWindow) {
		HWND hWndParent = NULL;
		m_spInPlaceSite->GetWindow(&hWndParent);
		if(::IsWindow(hWndParent)) {
			ATLVERIFY(RemoveWindowSubclass(hWndParent, WindowlessLabel::ParentWindowSubclass, reinterpret_cast<UINT_PTR>(this)));
		}
		flags.hasSubclassedParentWindow = FALSE;
	}
	return CComControl<WindowlessLabel>::IOleInPlaceObject_InPlaceDeactivate();
}

STDMETHODIMP WindowlessLabel::IOleInPlaceObject_SetObjectRects(LPCRECT pPositionRectangle, LPCRECT pClippingRectangle)
{
	BOOL resized = ((m_rcPos.right - m_rcPos.left != pPositionRectangle->right - pPositionRectangle->left) || (m_rcPos.bottom - m_rcPos.top != pPositionRectangle->bottom - pPositionRectangle->top));
	HRESULT hr = CComControl<WindowlessLabel>::IOleInPlaceObject_SetObjectRects(pPositionRectangle, pClippingRectangle);
	if(resized) {
		Raise_ResizedControlWindow();
	}
	return hr;
}

STDMETHODIMP WindowlessLabel::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

STDMETHODIMP WindowlessLabel::GetDropTarget(IDropTarget** ppDropTarget)
{
	if(!ppDropTarget) {
		return E_POINTER;
	}

	if(properties.registerForOLEDragDrop) {
		return QueryInterface(IID_PPV_ARGS(ppDropTarget));
	}
	return E_NOTIMPL;
}

BOOL WindowlessLabel::ClientToScreen(LPPOINT lpPoint) const
{
	if(lpPoint) {
		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		if(::IsWindow(hWndParent)) {
			lpPoint->x += m_rcPos.left;
			lpPoint->y += m_rcPos.top;
			::ClientToScreen(hWndParent, lpPoint);
			return TRUE;
		}
	}
	return FALSE;
}

BOOL WindowlessLabel::ClientToScreen(LPRECT lpRect) const
{
	if(lpRect) {
		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		if(::IsWindow(hWndParent)) {
			lpRect->left += m_rcPos.left;
			lpRect->top += m_rcPos.top;
			lpRect->right += m_rcPos.left;
			lpRect->bottom += m_rcPos.top;
			CWindow(hWndParent).ClientToScreen(lpRect);
			return TRUE;
		}
	}
	return FALSE;
}

BOOL WindowlessLabel::ScreenToClient(LPPOINT lpPoint) const
{
	if(lpPoint) {
		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		if(::IsWindow(hWndParent)) {
			::ScreenToClient(hWndParent, lpPoint);
			lpPoint->x -= m_rcPos.left;
			lpPoint->y -= m_rcPos.top;
			return TRUE;
		}
	}
	return FALSE;
}

BOOL WindowlessLabel::ScreenToClient(LPRECT lpRect) const
{
	if(lpRect) {
		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		if(::IsWindow(hWndParent)) {
			CWindow(hWndParent).ScreenToClient(lpRect);
			lpRect->left -= m_rcPos.left;
			lpRect->top -= m_rcPos.top;
			lpRect->right -= m_rcPos.left;
			lpRect->bottom -= m_rcPos.top;
			return TRUE;
		}
	}
	return FALSE;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP WindowlessLabel::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
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
		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		if(::IsWindow(hWndParent)) {
			dragDropStatus.pDropTargetHelper->DragEnter(hWndParent, pDataObject, &buffer, *pEffect);
		} else {
			dragDropStatus.pDropTargetHelper->Release();
			dragDropStatus.pDropTargetHelper = NULL;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
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

STDMETHODIMP WindowlessLabel::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
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
STDMETHODIMP WindowlessLabel::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
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

STDMETHODIMP WindowlessLabel::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_WLLBL_APPEARANCE:
		case DISPID_WLLBL_BACKSTYLE:
		case DISPID_WLLBL_BORDERSTYLE:
		case DISPID_WLLBL_CLIPLASTLINE:
		case DISPID_WLLBL_HALIGNMENT:
		case DISPID_WLLBL_MOUSEICON:
		case DISPID_WLLBL_MOUSEPOINTER:
		case DISPID_WLLBL_TEXTTRUNCATIONSTYLE:
		case DISPID_WLLBL_VALIGNMENT:
		case DISPID_WLLBL_WORDWRAPPING:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_WLLBL_DISABLEDEVENTS:
		case DISPID_WLLBL_DONTREDRAW:
		case DISPID_WLLBL_HOVERTIME:
		case DISPID_WLLBL_RIGHTTOLEFT:
		case DISPID_WLLBL_USEMNEMONIC:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_WLLBL_BACKCOLOR:
		case DISPID_WLLBL_FORECOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_WLLBL_APPID:
		case DISPID_WLLBL_APPNAME:
		case DISPID_WLLBL_APPSHORTNAME:
		case DISPID_WLLBL_BUILD:
		case DISPID_WLLBL_CHARSET:
		case DISPID_WLLBL_ISRELEASE:
		case DISPID_WLLBL_PROGRAMMER:
		case DISPID_WLLBL_TESTER:
		case DISPID_WLLBL_TEXT:
		case DISPID_WLLBL_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_WLLBL_REGISTERFOROLEDRAGDROP:
		case DISPID_WLLBL_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_WLLBL_FONT:
		case DISPID_WLLBL_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_WLLBL_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
		case DISPID_WLLBL_AUTOSIZE:
			*pCategory = PROPCAT_Position;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString WindowlessLabel::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString WindowlessLabel::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString WindowlessLabel::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=LabelControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString WindowlessLabel::GetSpecialThanks(void)
{
	return TEXT("Wine Headquarters");
}

CAtlString WindowlessLabel::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString WindowlessLabel::GetVersion(void)
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
STDMETHODIMP WindowlessLabel::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_WLLBL_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_WLLBL_AUTOSIZE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.autoSize), description);
			break;
		case DISPID_WLLBL_BACKSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.backStyle), description);
			break;
		case DISPID_WLLBL_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_WLLBL_HALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.hAlignment), description);
			break;
		case DISPID_WLLBL_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_WLLBL_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_WLLBL_TEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_WLLBL_TEXTTRUNCATIONSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.textTruncationStyle), description);
			break;
		case DISPID_WLLBL_VALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.vAlignment), description);
			break;
		case DISPID_WLLBL_WORDWRAPPING:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.wordWrapping), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<WindowlessLabel>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP WindowlessLabel::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_WLLBL_BACKSTYLE:
		case DISPID_WLLBL_BORDERSTYLE:
			c = 2;
			break;
		case DISPID_WLLBL_APPEARANCE:
		case DISPID_WLLBL_AUTOSIZE:
		case DISPID_WLLBL_HALIGNMENT:
		case DISPID_WLLBL_VALIGNMENT:
		case DISPID_WLLBL_WORDWRAPPING:
			c = 3;
			break;
		case DISPID_WLLBL_RIGHTTOLEFT:
		case DISPID_WLLBL_TEXTTRUNCATIONSTYLE:
			c = 4;
			break;
		case DISPID_WLLBL_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<WindowlessLabel>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_WLLBL_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
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

STDMETHODIMP WindowlessLabel::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_WLLBL_APPEARANCE:
		case DISPID_WLLBL_AUTOSIZE:
		case DISPID_WLLBL_BACKSTYLE:
		case DISPID_WLLBL_BORDERSTYLE:
		case DISPID_WLLBL_HALIGNMENT:
		case DISPID_WLLBL_MOUSEPOINTER:
		case DISPID_WLLBL_RIGHTTOLEFT:
		case DISPID_WLLBL_TEXTTRUNCATIONSTYLE:
		case DISPID_WLLBL_VALIGNMENT:
		case DISPID_WLLBL_WORDWRAPPING:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<WindowlessLabel>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	switch(property)
	{
		case DISPID_WLLBL_TEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
			break;
	}
	return IPerPropertyBrowsingImpl<WindowlessLabel>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT WindowlessLabel::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_WLLBL_APPEARANCE:
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
		case DISPID_WLLBL_AUTOSIZE:
			switch(cookie) {
				case asNone:
					description = GetResStringWithNumber(IDP_AUTOSIZENONE, asNone);
					break;
				case asGrowHorizontally:
					description = GetResStringWithNumber(IDP_AUTOSIZEGROWHORIZONTALLY, asGrowHorizontally);
					break;
				case asGrowVertically:
					description = GetResStringWithNumber(IDP_AUTOSIZEGROWVERTICALLY, asGrowVertically);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_WLLBL_BACKSTYLE:
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
		case DISPID_WLLBL_BORDERSTYLE:
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
		case DISPID_WLLBL_HALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HALIGNMENTLEFT, halLeft);
					break;
				case halCenter:
					description = GetResStringWithNumber(IDP_HALIGNMENTCENTER, halCenter);
					break;
				case halRight:
					description = GetResStringWithNumber(IDP_HALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_WLLBL_MOUSEPOINTER:
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
		case DISPID_WLLBL_RIGHTTOLEFT:
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
		case DISPID_WLLBL_TEXTTRUNCATIONSTYLE:
			switch(cookie) {
				case ttsNoEllipsis:
					description = GetResStringWithNumber(IDP_TEXTTRUNCATIONSTYLENOELLIPSIS, ttsNoEllipsis);
					break;
				case ttsEndEllipsis:
					description = GetResStringWithNumber(IDP_TEXTTRUNCATIONSTYLEENDELLIPSIS, ttsEndEllipsis);
					break;
				case ttsWordEllipsis:
					description = GetResStringWithNumber(IDP_TEXTTRUNCATIONSTYLEWORDELLIPSIS, ttsWordEllipsis);
					break;
				case ttsPathEllipsis:
					description = GetResStringWithNumber(IDP_TEXTTRUNCATIONSTYLEPATHELLIPSIS, ttsPathEllipsis);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_WLLBL_VALIGNMENT:
			switch(cookie) {
				case valTop:
					description = GetResStringWithNumber(IDP_VALIGNMENTTOP, valTop);
					break;
				case valCenter:
					description = GetResStringWithNumber(IDP_VALIGNMENTCENTER, valCenter);
					break;
				case valBottom:
					description = GetResStringWithNumber(IDP_VALIGNMENTBOTTOM, valBottom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_WLLBL_WORDWRAPPING:
			switch(cookie) {
				case wwManual:
					description = GetResStringWithNumber(IDP_WORDWRAPPINGMANUAL, wwManual);
					break;
				case wwAutomatic:
					description = GetResStringWithNumber(IDP_WORDWRAPPINGAUTOMATIC, wwAutomatic);
					break;
				case wwSingleLine:
					description = GetResStringWithNumber(IDP_WORDWRAPPINGSINGLELINE, wwSingleLine);
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
STDMETHODIMP WindowlessLabel::GetPages(CAUUID* pPropertyPages)
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
STDMETHODIMP WindowlessLabel::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			HRESULT hr = IOleObjectImpl<WindowlessLabel>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			if(DoesVerbActivate(verbID)) {
				if(m_spInPlaceSite && !flags.hasSubclassedParentWindow) {
					hWndParent = NULL;
					m_spInPlaceSite->GetWindow(&hWndParent);
					if(::IsWindow(hWndParent)) {
						flags.hasSubclassedParentWindow = SetWindowSubclass(hWndParent, WindowlessLabel::ParentWindowSubclass, reinterpret_cast<UINT_PTR>(this), NULL);
						ATLVERIFY(flags.hasSubclassedParentWindow);
					}
				}
			}
			return hr;
			break;
	}
}

STDMETHODIMP WindowlessLabel::EnumVerbs(IEnumOLEVERB** ppEnumerator)
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
STDMETHODIMP WindowlessLabel::GetControlInfo(LPCONTROLINFO pControlInfo)
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

STDMETHODIMP WindowlessLabel::OnMnemonic(LPMSG /*pMessage*/)
{
	if(!properties.enabled) {
		return S_OK;
	}

	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT WindowlessLabel::DoVerbAbout(HWND hWndParent)
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

HRESULT WindowlessLabel::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_WLLBL_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
				Redraw();
			}
			break;
	}
	return S_OK;
}

HRESULT WindowlessLabel::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP WindowlessLabel::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_APPEARANCE);
	if(newValue < 0 || newValue > a3DLight) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(properties.autoSize != asNone) {
			DoAutoSize();
		}
		Redraw();
		FireOnChanged(DISPID_WLLBL_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 17;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"LabelControls");
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"LblCtls");
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_AutoSize(AutoSizeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AutoSizeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.autoSize;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_AutoSize(AutoSizeConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_AUTOSIZE);
	if(properties.autoSize != newValue) {
		properties.autoSize = newValue;
		SetDirty(TRUE);

		if(properties.autoSize != asNone) {
			DoAutoSize();
		}
		FireOnChanged(DISPID_WLLBL_AUTOSIZE);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_BackStyle(BackStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BackStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.backStyle;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_BackStyle(BackStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_BACKSTYLE);
	if(properties.backStyle != newValue) {
		properties.backStyle = newValue;
		SetDirty(TRUE);

		if(m_spAdviseSink) {
			CComQIPtr<IAdviseSinkEx> pAdviseSinkEx = m_spClientSite;
			if(pAdviseSinkEx) {
				pAdviseSinkEx->OnViewStatusChange(_GetViewStatus());
			}
		}
		Redraw();
		FireOnChanged(DISPID_WLLBL_BACKSTYLE);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(properties.autoSize != asNone) {
			DoAutoSize();
		}
		Redraw();
		FireOnChanged(DISPID_WLLBL_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_CharSet(BSTR* pValue)
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

STDMETHODIMP WindowlessLabel::get_ClipLastLine(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.clipLastLine);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_ClipLastLine(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_CLIPLASTLINE);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.clipLastLine != b) {
		properties.clipLastLine = b;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_CLIPLASTLINE);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_WLLBL_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		Redraw();
		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_WLLBL_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Font(IFontDisp** ppFont)
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

STDMETHODIMP WindowlessLabel::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_WLLBL_FONT);
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
		Redraw();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_WLLBL_FONT);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_WLLBL_FONT);
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
		Redraw();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_WLLBL_FONT);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_HAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hAlignment;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_HAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_HALIGNMENT);
	if(properties.hAlignment != newValue) {
		properties.hAlignment = newValue;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_HALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_WLLBL_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP WindowlessLabel::get_MouseIcon(IPictureDisp** ppMouseIcon)
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

STDMETHODIMP WindowlessLabel::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_WLLBL_MOUSEICON);
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
	FireOnChanged(DISPID_WLLBL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_WLLBL_MOUSEICON);
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
	FireOnChanged(DISPID_WLLBL_MOUSEICON);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_WLLBL_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_REGISTERFOROLEDRAGDROP);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_WLLBL_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_WLLBL_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_Text(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_TEXT);
	properties.text.AssignBSTR(newValue);
	SetDirty(TRUE);

	// update the accelerator table
	TCHAR newAccelerator = 0;
	for(long i = properties.text.Length() - 1; i > 0; --i) {
		if((properties.text[i - 1] == L'&') && (properties.text[i] != L'&')) {
			COLE2T converter(properties.text);
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

	if(properties.autoSize != asNone) {
		DoAutoSize();
	}
	Redraw();

	if(!(properties.disabledEvents & deTextChangedEvents)) {
		Raise_TextChanged();
	}
	FireOnChanged(DISPID_WLLBL_TEXT);
	SendOnDataChange();
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_TextTruncationStyle(TextTruncationStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, TextTruncationStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.textTruncationStyle;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_TextTruncationStyle(TextTruncationStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_TEXTTRUNCATIONSTYLE);
	if(properties.textTruncationStyle != newValue) {
		properties.textTruncationStyle = newValue;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_TEXTTRUNCATIONSTYLE);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_UseMnemonic(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useMnemonic);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_UseMnemonic(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_USEMNEMONIC);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useMnemonic != b) {
		properties.useMnemonic = b;
		SetDirty(TRUE);

		if(properties.autoSize != asNone) {
			DoAutoSize();
		}
		Redraw();
		FireOnChanged(DISPID_WLLBL_USEMNEMONIC);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		ApplyFont();
		Redraw();
		FireOnChanged(DISPID_WLLBL_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_VAlignment(VAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, VAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.vAlignment;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_VAlignment(VAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_VALIGNMENT);
	if(properties.vAlignment != newValue) {
		properties.vAlignment = newValue;
		SetDirty(TRUE);

		Redraw();
		FireOnChanged(DISPID_WLLBL_VALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::get_Version(BSTR* pValue)
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

STDMETHODIMP WindowlessLabel::get_WordWrapping(WordWrappingConstants* pValue)
{
	ATLASSERT_POINTER(pValue, WordWrappingConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.wordWrapping;
	return S_OK;
}

STDMETHODIMP WindowlessLabel::put_WordWrapping(WordWrappingConstants newValue)
{
	PUTPROPPROLOG(DISPID_WLLBL_WORDWRAPPING);
	if(properties.wordWrapping != newValue) {
		properties.wordWrapping = newValue;
		SetDirty(TRUE);

		if(properties.autoSize != asNone) {
			DoAutoSize();
		}
		Redraw();
		FireOnChanged(DISPID_WLLBL_WORDWRAPPING);
	}
	return S_OK;
}

STDMETHODIMP WindowlessLabel::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP WindowlessLabel::FinishOLEDragDrop(void)
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

STDMETHODIMP WindowlessLabel::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP WindowlessLabel::Refresh(void)
{
	Redraw();
	return S_OK;
}

STDMETHODIMP WindowlessLabel::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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


LRESULT WindowlessLabel::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		mousePosition.x -= m_rcPos.left;
		mousePosition.y -= m_rcPos.top;
		Raise_DblClick(button, shift, mousePosition.x, mousePosition.y);
	}

	return 0;
}

LRESULT WindowlessLabel::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		mousePosition.x -= m_rcPos.left;
		mousePosition.y -= m_rcPos.top;
		Raise_MDblClick(button, shift, mousePosition.x, mousePosition.y);
	}

	return 0;
}

LRESULT WindowlessLabel::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	mouseStatus.lastPosition = mousePosition;
	mightHaveLeftControl = FALSE;
	Raise_MouseMove(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		mousePosition.x -= m_rcPos.left;
		mousePosition.y -= m_rcPos.top;
		Raise_RDblClick(button, shift, mousePosition.x, mousePosition.y);
	}

	return 0;
}

LRESULT WindowlessLabel::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnSetCursor(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = FALSE;
	wasHandled = FALSE;
	BOOL setCursor = TRUE;

	if(m_spInPlaceSite) {
		// this will reset the cursor to the default
		HRESULT hr = m_spInPlaceSite->OnDefWindowMessage(message, wParam, lParam, &lr);
		if(SUCCEEDED(hr)) {
			wasHandled = TRUE;
			//setCursor = (hr == S_FALSE);
		} else {
			lr = FALSE;
		}
	}

	if(setCursor) {
		HCURSOR hCursor = NULL;
		if(properties.mousePointer == mpCustom) {
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
			lr = TRUE;
		}
	}
	return lr;
}

LRESULT WindowlessLabel::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Redraw();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT WindowlessLabel::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
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
		POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		mousePosition.x -= m_rcPos.left;
		mousePosition.y -= m_rcPos.top;
		Raise_XDblClick(button, shift, mousePosition.x, mousePosition.y);
	}

	return 0;
}

LRESULT WindowlessLabel::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}

LRESULT WindowlessLabel::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
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
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	mousePosition.x -= m_rcPos.left;
	mousePosition.y -= m_rcPos.top;
	Raise_MouseUp(button, shift, mousePosition.x, mousePosition.y);

	return 0;
}


void WindowlessLabel::ApplyFont(void)
{
	if(!properties.font.owningFontResource) {
		properties.font.currentFont.Detach();
	}
	properties.font.currentFont.Attach(NULL);

	if(properties.useSystemFont) {
		// use the system font
		properties.font.currentFont.Attach(static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));
		properties.font.owningFontResource = FALSE;
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
			}
		}
	}
	//Redraw();					// we call ApplyFont from OnDraw!
}

void WindowlessLabel::DoAutoSize(void)
{
	HDC hCompatibleDC = ::GetDC(HWND_DESKTOP);
	if(hCompatibleDC) {
		HDC hMemoryDC = CreateCompatibleDC(hCompatibleDC);
		if(hMemoryDC) {
			if(properties.font.currentFont.IsNull()) {
				ApplyFont();
			}
			HGDIOBJ hPreviousFont = SelectObject(hMemoryDC, properties.font.currentFont);

			DWORD drawTextFlags = DT_CALCRECT | DT_LEFT | DT_TOP;
			if(properties.autoSize == asGrowHorizontally) {
				switch(properties.wordWrapping) {
					case wwSingleLine:
						drawTextFlags |= DT_SINGLELINE;
						break;
				}
			} else if(properties.autoSize == asGrowVertically) {
				switch(properties.wordWrapping) {
					case wwAutomatic:
						drawTextFlags |= DT_WORDBREAK;
						break;
					case wwSingleLine:
						drawTextFlags |= DT_SINGLELINE;
						break;
				}
			}
			if(!properties.clipLastLine) {
				drawTextFlags |= DT_EDITCONTROL;
			}
			switch(properties.textTruncationStyle) {
				case ttsEndEllipsis:
					drawTextFlags |= DT_END_ELLIPSIS;
					break;
				case ttsWordEllipsis:
					drawTextFlags |= DT_WORD_ELLIPSIS;
					break;
				case ttsPathEllipsis:
					drawTextFlags |= DT_PATH_ELLIPSIS;
					break;
			}
			if(!properties.useMnemonic) {
				drawTextFlags |= DT_NOPREFIX;
			}

			CRect windowRectangle = m_rcPos;
			CRect contentRectangle;
			if(properties.autoSize == asGrowVertically) {
				contentRectangle.left = 0;
				contentRectangle.right = windowRectangle.Width();
			}
			COLE2T converter(properties.text);
			DrawText(hMemoryDC, converter, -1, &contentRectangle, drawTextFlags);

			DWORD windowStyle = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
			switch(properties.borderStyle) {
				case bsFixedSingle:
					windowStyle |= WS_BORDER;
					break;
			}
			DWORD windowExStyle = 0;
			switch(properties.appearance) {
				case a3D:
					windowExStyle |= WS_EX_CLIENTEDGE;
					break;
				case a3DLight:
					windowExStyle |= WS_EX_STATICEDGE;
					break;
			}
			AdjustWindowRectEx(&contentRectangle, windowStyle, FALSE, windowExStyle);

			if(windowRectangle.Width() != contentRectangle.Width() || windowRectangle.Height() != contentRectangle.Height()) {
				HAlignmentConstants hAlignment = halLeft;
				get_HAlignment(&hAlignment);
				if(properties.rightToLeft & rtlLayout) {
					switch(hAlignment) {
						case halLeft:
							hAlignment = halRight;
							break;
						case halRight:
							hAlignment = halLeft;
							break;
					}
				}
				switch(hAlignment) {
					case halLeft:
						windowRectangle.right = windowRectangle.left + contentRectangle.Width();
						break;
					case halCenter:
						windowRectangle.left -= (contentRectangle.Width() - windowRectangle.Width()) / 2;			// don't bit-shift here because the difference can be negative
						windowRectangle.right = windowRectangle.left + contentRectangle.Width();
						break;
					case halRight:
						windowRectangle.left = windowRectangle.right - contentRectangle.Width();
						break;
				}
				VAlignmentConstants vAlignment = valTop;
				get_VAlignment(&vAlignment);
				switch(vAlignment) {
					case valTop:
						windowRectangle.bottom = windowRectangle.top + contentRectangle.Height();
						break;
					case valCenter:
						windowRectangle.top -= (contentRectangle.Height() - windowRectangle.Height()) / 2;			// don't bit-shift here because the difference can be negative
						windowRectangle.bottom = windowRectangle.top + contentRectangle.Height();
						break;
					case valBottom:
						windowRectangle.top = windowRectangle.bottom - contentRectangle.Height();
						break;
				}

				if(m_spInPlaceSite) {
					m_spInPlaceSite->OnPosRectChange(&windowRectangle);
				}
			}

			SelectObject(hMemoryDC, hPreviousFont);
			DeleteDC(hMemoryDC);
		}
		::ReleaseDC(HWND_DESKTOP, hCompatibleDC);
	}
}


inline HRESULT WindowlessLabel::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Click(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		if(x == -1 && y == -1) {
			// the event was caused by the keyboard
			/*if(properties.processContextMenuKeys) {
				// propose the middle of the control's client rectangle as the menu's position
				CRect clientRectangle;
				GetClientRect(&clientRectangle);
				CPoint centerPoint = clientRectangle.CenterPoint();
				x = centerPoint.x;
				y = centerPoint.y;
			} else {*/
				return S_OK;
			//}
		}

		return Fire_ContextMenu(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL fireMouseDown = FALSE;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
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
	if(m_spInPlaceSite) {
		if(SUCCEEDED(m_spInPlaceSite->SetCapture(TRUE))) {
			// As long as we hold the capture, we won't receive any WM_SETCURSOR messages.
			BOOL dummy = TRUE;
			OnSetCursor(WM_SETCURSOR, MAKELPARAM(HTCLIENT, WM_LBUTTONDOWN), 0, dummy);
		}
	}

	HRESULT hr = S_OK;
	if(fireMouseDown) {
		hr = Fire_MouseDown(button, shift, x, y);
	}
	return hr;
}

inline HRESULT WindowlessLabel::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hoverTime = 0;
	if(properties.hoverTime == -1) {
		SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, &hoverTime, 0);
	} else {
		hoverTime = properties.hoverTime;
	}
	if(hoverTime > 0) {
		if(mouseHoverTimerID) {
			::KillTimer(NULL, mouseHoverTimerID);
			EnterCriticalSection(&timerMapCriticalSection);
			#ifdef USE_STL
				std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(mouseHoverTimerID);
				if(it != timerMap.end()) {
					timerMap.erase(it);
				}
			#else
				timerMap.RemoveKey(mouseHoverTimerID);
			#endif
			LeaveCriticalSection(&timerMapCriticalSection);
			mouseHoverTimerID = 0;
		}
		mouseStatus.hoverStartPosition.x = x;
		mouseStatus.hoverStartPosition.y = y;
		mouseHoverTimerID = ::SetTimer(NULL, 0, hoverTime, MouseHoverTimerProc);
		if(mouseHoverTimerID) {
			EnterCriticalSection(&timerMapCriticalSection);
			timerMap[mouseHoverTimerID] = this;
			LeaveCriticalSection(&timerMapCriticalSection);
		}
	}

	if(TIMERINT_MOUSELEAVE > 0) {
		if(mouseLeaveTimerID) {
			::KillTimer(NULL, mouseLeaveTimerID);
			EnterCriticalSection(&timerMapCriticalSection);
			#ifdef USE_STL
				std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(mouseLeaveTimerID);
				if(it != timerMap.end()) {
					timerMap.erase(it);
				}
			#else
				timerMap.RemoveKey(mouseLeaveTimerID);
			#endif
			LeaveCriticalSection(&timerMapCriticalSection);
			mouseLeaveTimerID = 0;
		}
		mouseLeaveTimerID = ::SetTimer(NULL, 0, TIMERINT_MOUSELEAVE, MouseLeaveTimerProc);
		if(mouseLeaveTimerID) {
			EnterCriticalSection(&timerMapCriticalSection);
			timerMap[mouseLeaveTimerID] = this;
			LeaveCriticalSection(&timerMapCriticalSection);
		}
	}

	mouseStatus.EnterControl();

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseEnter(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.hoveredControl) {
		if(mouseHoverTimerID) {
			::KillTimer(NULL, mouseHoverTimerID);
			EnterCriticalSection(&timerMapCriticalSection);
			#ifdef USE_STL
				std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(mouseHoverTimerID);
				if(it != timerMap.end()) {
					timerMap.erase(it);
				}
			#else
				timerMap.RemoveKey(mouseHoverTimerID);
			#endif
			LeaveCriticalSection(&timerMapCriticalSection);
			mouseHoverTimerID = 0;
		}
		mouseStatus.HoverControl();

		if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
			return Fire_MouseHover(button, shift, x, y);
		}
	}
	return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(m_spInPlaceSite) {
		if(m_spInPlaceSite->GetCapture() == S_OK) {
			m_spInPlaceSite->SetCapture(FALSE);
		}
	}
	if(mouseHoverTimerID) {
		::KillTimer(NULL, mouseHoverTimerID);
		EnterCriticalSection(&timerMapCriticalSection);
		#ifdef USE_STL
			std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(mouseHoverTimerID);
			if(it != timerMap.end()) {
				timerMap.erase(it);
			}
		#else
			timerMap.RemoveKey(mouseHoverTimerID);
		#endif
		LeaveCriticalSection(&timerMapCriticalSection);
		mouseHoverTimerID = 0;
	}
	if(mouseLeaveTimerID) {
		::KillTimer(NULL, mouseLeaveTimerID);
		EnterCriticalSection(&timerMapCriticalSection);
		#ifdef USE_STL
			std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(mouseLeaveTimerID);
			if(it != timerMap.end()) {
				timerMap.erase(it);
			}
		#else
			timerMap.RemoveKey(mouseLeaveTimerID);
		#endif
		LeaveCriticalSection(&timerMapCriticalSection);
		mouseLeaveTimerID = 0;
	}
	mouseStatus.LeaveControl();
	mightHaveLeftControl = FALSE;

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseLeave(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL hasLeftControl = FALSE;
	POINT cursorPosition = {x, y};
	ClientToScreen(&cursorPosition);
	RECT clientArea = m_rcPos;
	if(m_spInPlaceSite) {
		// exclude overlapped parts
		m_spInPlaceSite->AdjustRect(&clientArea);
	}
	OffsetRect(&clientArea, -m_rcPos.left, -m_rcPos.top);
	ClientToScreen(&clientArea);
	if(PtInRect(&clientArea, cursorPosition)) {
		HWND hWndParent = NULL;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->GetWindow(&hWndParent);
		}
		if(WindowFromPoint(cursorPosition) != hWndParent) {
			hasLeftControl = TRUE;
		}
	} else {
		hasLeftControl = TRUE;
	}

	if(hasLeftControl && !mouseStatus.IsClickCandidate(button)) {
		return Raise_MouseLeave(button, shift, x, y);
	}

	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}

	if(!mouseStatus.hoveredControl) {
		LONG mouseHoverHeight = 0;
		LONG mouseHoverWidth = 0;
		SystemParametersInfo(SPI_GETMOUSEHOVERHEIGHT, 0, &mouseHoverHeight, 0);
		SystemParametersInfo(SPI_GETMOUSEHOVERWIDTH, 0, &mouseHoverWidth, 0);
		RECT mouseHoverRect = {mouseStatus.hoverStartPosition.x - (mouseHoverWidth >> 1), mouseStatus.hoverStartPosition.y - (mouseHoverHeight >> 1), mouseStatus.hoverStartPosition.x + (mouseHoverWidth >> 1), mouseStatus.hoverStartPosition.y + (mouseHoverHeight >> 1)};
		if(!PtInRect(&mouseHoverRect, mouseStatus.lastPosition)) {
			// we're no longer within the hover rectangle - restart the timer
			UINT hoverTime = 0;
			if(properties.hoverTime == -1) {
				SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, &hoverTime, 0);
			} else {
				hoverTime = properties.hoverTime;
			}
			if(hoverTime > 0) {
				if(mouseHoverTimerID) {
					::KillTimer(NULL, mouseHoverTimerID);
					EnterCriticalSection(&timerMapCriticalSection);
					#ifdef USE_STL
						std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(mouseHoverTimerID);
						if(it != timerMap.end()) {
							timerMap.erase(it);
						}
					#else
						timerMap.RemoveKey(mouseHoverTimerID);
					#endif
					LeaveCriticalSection(&timerMapCriticalSection);
					mouseHoverTimerID = 0;
				}
				mouseStatus.hoverStartPosition = mouseStatus.lastPosition;
				mouseHoverTimerID = ::SetTimer(NULL, 0, hoverTime, MouseHoverTimerProc);
				if(mouseHoverTimerID) {
					EnterCriticalSection(&timerMapCriticalSection);
					timerMap[mouseHoverTimerID] = this;
					LeaveCriticalSection(&timerMapCriticalSection);
				}
			}
		}
	}

	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		return Fire_MouseMove(button, shift, x, y);
	}
	return S_OK;
}

inline HRESULT WindowlessLabel::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(button, shift, x, y);
	}

	if(m_spInPlaceSite) {
		if(m_spInPlaceSite->GetCapture() == S_OK) {
			m_spInPlaceSite->SetCapture(FALSE);
		}
	}
	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		POINT cursorPosition = {x, y};
		ClientToScreen(&cursorPosition);
		RECT clientArea = m_rcPos;
		if(m_spInPlaceSite) {
			// exclude overlapped parts
			m_spInPlaceSite->AdjustRect(&clientArea);
		}
		OffsetRect(&clientArea, -m_rcPos.left, -m_rcPos.top);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			HWND hWndParent = NULL;
			if(m_spInPlaceSite) {
				m_spInPlaceSite->GetWindow(&hWndParent);
			}
			if(WindowFromPoint(cursorPosition) != hWndParent) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
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
					Raise_ContextMenu(button, shift, x, y);
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
		} else {
			Raise_MouseLeave(button, shift, x, y);
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT WindowlessLabel::Raise_OLEDragDrop(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

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
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();

	return hr;
}

inline HRESULT WindowlessLabel::Raise_OLEDragEnter(IDataObject* pData, LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	return hr;
}

inline HRESULT WindowlessLabel::Raise_OLEDragLeave()
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y);
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();

	return hr;
}

inline HRESULT WindowlessLabel::Raise_OLEDragMouseMove(LPDWORD pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), button, shift, mousePosition.x, mousePosition.y);
		}
	//}

	return hr;
}

inline HRESULT WindowlessLabel::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RDblClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_TextChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TextChanged();
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XClick(button, shift, x, y);
	//}
	//return S_OK;
}

inline HRESULT WindowlessLabel::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_XDblClick(button, shift, x, y);
	//}
	//return S_OK;
}


void WindowlessLabel::RecreateControlWindow(void)
{
	if(m_bInPlaceActive && !flags.dontRecreate) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

HCURSOR WindowlessLabel::MousePointerConst2hCursor(MousePointerConstants mousePointer)
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

LRESULT CALLBACK WindowlessLabel::ParentWindowSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/)
{
	LRESULT lr = 0;
	WindowlessLabel* pThis = reinterpret_cast<WindowlessLabel*>(idSubclass);
	if(pThis) {
		switch(message) {
			case WM_XBUTTONDBLCLK:
			case WM_XBUTTONDOWN:
			case WM_XBUTTONUP:
			{
				BOOL outsideControl = FALSE;
				POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
				HWND hWndParent = NULL;
				if(pThis->m_spInPlaceSite) {
					pThis->m_spInPlaceSite->GetWindow(&hWndParent);
				}
				if(hWndParent) {
					::ClientToScreen(hWndParent, &mousePosition);
					RECT clientArea = pThis->m_rcPos;
					if(pThis->m_spInPlaceSite) {
						// exclude overlapped parts
						pThis->m_spInPlaceSite->AdjustRect(&clientArea);
					}
					OffsetRect(&clientArea, -pThis->m_rcPos.left, -pThis->m_rcPos.top);
					pThis->ClientToScreen(&clientArea);
					if(PtInRect(&clientArea, mousePosition)) {
						if(WindowFromPoint(mousePosition) != hWndParent) {
							outsideControl = TRUE;
						}
					} else {
						outsideControl = TRUE;
					}
				} else {
					outsideControl = TRUE;
				}

				if(!outsideControl) {
					if(pThis->ProcessWindowMessage(hWnd, message, wParam, lParam, lr, 0)) {
						return lr;
					}
				}
				break;
			}
		}
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
		switch(message) {
			case WM_CHANGEUISTATE:
				pThis->Redraw();
				break;
		}
	} else {
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
	}
	return lr;
}

void WindowlessLabel::MouseHoverTimerProc(HWND /*hWnd*/, UINT /*message*/, UINT_PTR timerID, DWORD /*time*/)
{
	WindowlessLabel* pControl = NULL;
	EnterCriticalSection(&timerMapCriticalSection);
	#ifdef USE_STL
		std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(timerID);
		if(it != timerMap.end()) {
			pControl = it->second;
		}
	#else
		CAtlMap<UINT_PTR, WindowlessLabel*>::CPair* pPair = timerMap.Lookup(timerID);
		if(pPair) {
			pControl = pPair->m_value;
		}
	#endif
	LeaveCriticalSection(&timerMapCriticalSection);

	if(pControl) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		pControl->Raise_MouseHover(button, shift, pControl->mouseStatus.lastPosition.x, pControl->mouseStatus.lastPosition.y);
	}
}

void WindowlessLabel::MouseLeaveTimerProc(HWND /*hWnd*/, UINT /*message*/, UINT_PTR timerID, DWORD /*time*/)
{
	WindowlessLabel* pControl = NULL;
	EnterCriticalSection(&timerMapCriticalSection);
	#ifdef USE_STL
		std::map<UINT_PTR, WindowlessLabel*>::iterator it = timerMap.find(timerID);
		if(it != timerMap.end()) {
			pControl = it->second;
		}
	#else
		CAtlMap<UINT_PTR, WindowlessLabel*>::CPair* pPair = timerMap.Lookup(timerID);
		if(pPair) {
			pControl = pPair->m_value;
		}
	#endif
	LeaveCriticalSection(&timerMapCriticalSection);

	if(pControl) {
		BOOL hasLeftControl = FALSE;
		POINT cursorPosition = {0};
		GetCursorPos(&cursorPosition);
		pControl->ScreenToClient(&cursorPosition);
		if(cursorPosition.x != pControl->mouseStatus.lastPosition.x || cursorPosition.y != pControl->mouseStatus.lastPosition.y) {
			/* The mouse cursor has moved, but we did not receive a mouse move message, so this means we're no
			   longer below the cursor. It could also mean that the mouse has been moved very fast and the timer
			   fired before the mouse move event. Therefore we give it a second chance. */
			if(pControl->mightHaveLeftControl) {
				hasLeftControl = TRUE;
			} else {
				pControl->mightHaveLeftControl = TRUE;
			}
		}

		if(hasLeftControl) {
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			if(!pControl->mouseStatus.IsClickCandidate(button)) {
				pControl->Raise_MouseLeave(button, shift, pControl->mouseStatus.lastPosition.x, pControl->mouseStatus.lastPosition.y);
			}
		}
	}
}

BOOL WindowlessLabel::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}