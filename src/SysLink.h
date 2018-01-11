//////////////////////////////////////////////////////////////////////
/// \class SysLink
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c SysLink</em>
///
/// This class superclasses \c SysLink and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Buffer the visited link color?
///
/// \sa LblCtlsLibU::ISysLink
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#include "LblCtlsU.h"
#include "_ISysLinkEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "Link.h"
#include "Links.h"
#include "CommonProperties.h"
#include "StringProperties.h"
#include "TargetOLEDataObject.h"


class ATL_NO_VTABLE SysLink : 
    public CComObjectRootEx<CComSingleThreadModel>,
   	public IDispatchImpl<ISysLink, &IID_ISysLink, &LIBID_LblCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    public IPersistStreamInitImpl<SysLink>,
    public IOleControlImpl<SysLink>,
    public IOleObjectImpl<SysLink>,
    public IOleInPlaceActiveObjectImpl<SysLink>,
    public IViewObjectExImpl<SysLink>,
    public IOleInPlaceObjectWindowlessImpl<SysLink>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<SysLink>,
    public Proxy_ISysLinkEvents<SysLink>,
    public IPersistStorageImpl<SysLink>,
    public IPersistPropertyBagImpl<SysLink>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<SysLink>,
   	public IProvideClassInfo2Impl<&CLSID_SysLink, &__uuidof(_ISysLinkEvents), &LIBID_LblCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    public IPropertyNotifySinkCP<SysLink>,
    public CComCoClass<SysLink, &CLSID_SysLink>,
    public CComControl<SysLink>,
    public IPerPropertyBrowsingImpl<SysLink>,
    public IDropTarget,
    public ICategorizeProperties,
    public ICreditsProvider
{
	friend class Link;
	friend class Links;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	SysLink();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_SYSLINK)

		DECLARE_WND_SUPERCLASS(TEXT("SysLinkU"), WC_LINK)

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		DWORD _GetViewStatus();
		/*// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)*/

		BEGIN_COM_MAP(SysLink)
			COM_INTERFACE_ENTRY(ISysLink)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
		END_COM_MAP()

		BEGIN_PROP_MAP(SysLink)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("Appearance", DISPID_SL_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutomaticallyMarkLinksAsVisited", DISPID_SL_AUTOMATICALLYMARKLINKSASVISITED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("BackColor", DISPID_SL_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BackStyle", DISPID_SL_BACKSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_SL_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DetectDoubleClicks", DISPID_SL_DETECTDOUBLECLICKS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_SL_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_SL_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Enabled", DISPID_SL_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_SL_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_SL_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HAlignment", DISPID_SL_HALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HotTracking", DISPID_SL_HOTTRACKING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("HoverTime", DISPID_SL_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_SL_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_SL_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_SL_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ProcessReturnKey", DISPID_SL_PROCESSRETURNKEY, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_SL_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_SL_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ShowToolTips", DISPID_SL_SHOWTOOLTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_SL_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			//PROP_ENTRY_TYPE("Text", DISPID_SL_TEXT, CLSID_StringProperties, VT_BSTR)
			PROP_ENTRY_TYPE("UseMnemonic", DISPID_SL_USEMNEMONIC, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_SL_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseVisualStyle", DISPID_SL_USEVISUALSTYLE, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(SysLink)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_ISysLinkEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(SysLink)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_NCHITTEST, OnNCHitTest)
			MESSAGE_HANDLER(WM_NCMOUSEMOVE, OnNCMouseMove)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_PRINTCLIENT, OnPaint)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTEXT, OnSetText)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_SIZE, OnSize)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM_CTLCOLORSTATIC, OnReflectedCtlColorStatic)
			MESSAGE_HANDLER(OCM_NOTIFY, OnReflectedNotify)

			MESSAGE_HANDLER(LM_SETITEM, OnSetItem)

			REFLECTED_NOTIFY_CODE_HANDLER(NM_CLICK, OnClickNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(NM_CUSTOMTEXT, OnCustomTextNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(NM_RETURN, OnReturnNotification)

			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOA, OnToolTipGetDispInfoNotificationA)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOW, OnToolTipGetDispInfoNotificationW)

			CHAIN_MSG_MAP(CComControl<SysLink>)
		END_MSG_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist a Unicode text property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISysLink
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Appearance, get_BorderStyle, LblCtlsLibU::AppearanceConstants
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Appearance, put_BorderStyle, LblCtlsLibU::AppearanceConstants
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c AutomaticallyMarkLinksAsVisited property</em>
	///
	/// Retrieves whether a link gets marked as visited if it is opened. If set to \c VARIANT_TRUE, a link
	/// becomes marked as visited if the \c OpenLink event is raised for it; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE, the \c CustomDraw event will be fired for each
	///          single link.
	///
	/// \sa put_AutomaticallyMarkLinksAsVisited, get_HotTracking, Link::get_Visited, Raise_OpenLink
	virtual HRESULT STDMETHODCALLTYPE get_AutomaticallyMarkLinksAsVisited(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutomaticallyMarkLinksAsVisited property</em>
	///
	/// Sets whether a link gets marked as visited if it is opened. If set to \c VARIANT_TRUE, a link
	/// becomes marked as visited if the \c OpenLink event is raised for it; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c VARIANT_TRUE, the \c CustomDraw event will be fired for each
	///          single link.
	///
	/// \sa get_AutomaticallyMarkLinksAsVisited, put_HotTracking, Link::put_Visited, Raise_OpenLink
	virtual HRESULT STDMETHODCALLTYPE put_AutomaticallyMarkLinksAsVisited(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed link controls.
	///
	/// \sa put_BackColor, get_ForeColor, get_BackStyle, get_UseVisualStyle
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed link controls.
	///
	/// \sa get_BackColor, put_ForeColor, put_BackStyle, put_UseVisualStyle
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BackStyle property</em>
	///
	/// Retrieves how the control's background is drawn. Any of the values defined by the
	/// \c BackStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BackStyle, get_BackColor, LblCtlsLibU::BackStyleConstants
	/// \else
	///   \sa put_BackStyle, get_BackColor, LblCtlsLibA::BackStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BackStyle(BackStyleConstants* pValue);
	/// \brief <em>Sets the \c BackStyle property</em>
	///
	/// Sets how the control's background is drawn. Any of the values defined by the
	/// \c BackStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BackStyle, put_BackColor, LblCtlsLibU::BackStyleConstants
	/// \else
	///   \sa get_BackStyle, put_BackColor, LblCtlsLibA::BackStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BackStyle(BackStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BorderStyle, get_Appearance, LblCtlsLibU::BorderStyleConstants
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BorderStyle, put_Appearance, LblCtlsLibU::BorderStyleConstants
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c CaretLink property</em>
	///
	/// Retrieves the control's caret link. The caret link is the link that has the focus.
	///
	/// \param[out] ppCaretLink Receives the caret link's \c ILink implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa putref_CaretLink, Link::get_Caret
	virtual HRESULT STDMETHODCALLTYPE get_CaretLink(ILink** ppCaretLink);
	/// \brief <em>Sets the \c CaretLink property</em>
	///
	/// Sets the control's caret link. The caret link is the link that has the focus.
	///
	/// \param[in] pNewCaretLink The new caret link's \c ILink implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaretLink
	virtual HRESULT STDMETHODCALLTYPE putref_CaretLink(ILink* pNewCaretLink);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c DetectDoubleClicks property</em>
	///
	/// Retrieves whether double clicks are enabled or disabled. If set to \c VARIANT_TRUE, double clicks are
	/// accepted; otherwise all clicks are handled as single clicks.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DetectDoubleClicks, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	virtual HRESULT STDMETHODCALLTYPE get_DetectDoubleClicks(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DetectDoubleClicks property</em>
	///
	/// Enables or disables double clicks. If set to \c VARIANT_TRUE, double clicks are accepted; otherwise
	/// all clicks are handled as single clicks.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DetectDoubleClicks, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	virtual HRESULT STDMETHODCALLTYPE put_DetectDoubleClicks(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DisabledEvents, LblCtlsLibU::DisabledEventsConstants
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DisabledEvents, LblCtlsLibU::DisabledEventsConstants
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	///
	/// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	/// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DontRedraw property</em>
	///
	/// Enables or disables automatic redrawing of the control. Disabling redraw while doing large changes
	/// on the control may increase performance. If set to \c VARIANT_FALSE, the control will redraw itself
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, Link::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, Link::put_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the control's content.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont, put_ForeColor, get_UseVisualStyle,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont, put_ForeColor, put_UseVisualStyle,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the control's content.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont, put_ForeColor, put_UseVisualStyle,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the control's text color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ForeColor, get_BackColor, get_UseVisualStyle
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ForeColor, put_BackColor, put_UseVisualStyle
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c HAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the control's content. Some of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HAlignment, get_Text, LblCtlsLibU::HAlignmentConstants
	virtual HRESULT STDMETHODCALLTYPE get_HAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c HAlignment property</em>
	///
	/// Sets the horizontal alignment of the control's content. Some of the values defined by the
	/// \c HAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_HAlignment, put_Text, LblCtlsLibU::HAlignmentConstants
	virtual HRESULT STDMETHODCALLTYPE put_HAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HotTracking property</em>
	///
	/// Retrieves whether the link underneath the mouse cursor becomes highlighted. If set to
	/// \c VARIANT_TRUE, the link becomes highlighted; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HotTracking, get_AutomaticallyMarkLinksAsVisited, Link::get_Hot
	virtual HRESULT STDMETHODCALLTYPE get_HotTracking(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HotTracking property</em>
	///
	/// Sets whether the link underneath the mouse cursor becomes highlighted. If set to
	/// \c VARIANT_TRUE, the link becomes highlighted; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HotTracking, put_AutomaticallyMarkLinksAsVisited, Link::put_Hot
	virtual HRESULT STDMETHODCALLTYPE put_HotTracking(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndToolTip property</em>
	///
	/// Retrieves the tooltip control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndToolTip, get_ShowToolTips
	virtual HRESULT STDMETHODCALLTYPE get_hWndToolTip(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndToolTip property</em>
	///
	/// Sets the window handle of the tooltip control to use.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set tooltip window does NOT get destroyed automatically.
	///
	/// \sa get_hWndToolTip, put_ShowToolTips
	virtual HRESULT STDMETHODCALLTYPE put_hWndToolTip(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Links property</em>
	///
	/// Retrieves a collection object wrapping the control's hyper links.
	///
	/// \param[out] ppLinks Receives the collection object's \c ILinks implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Links
	virtual HRESULT STDMETHODCALLTYPE get_Links(ILinks** ppLinks);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, LblCtlsLibU::MousePointerConstants,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, LblCtlsLibU::MousePointerConstants,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MouseIcon, put_MouseIcon, put_MousePointer, LblCtlsLibU::MousePointerConstants,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MousePointer, get_MouseIcon, LblCtlsLibU::MousePointerConstants
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MousePointer, put_MouseIcon, LblCtlsLibU::MousePointerConstants
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10]
	/// or [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the events are fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessReturnKey property</em>
	///
	/// Retrieves whether the control fires the \c OpenLink event if the user presses the \c [RETURN]
	/// key. If set to \c VARIANT_TRUE, the event gets fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessReturnKey, Raise_OpenLink
	virtual HRESULT STDMETHODCALLTYPE get_ProcessReturnKey(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessReturnKey property</em>
	///
	/// Sets whether the control fires the \c OpenLink event if the user presses the \c [RETURN]
	/// key. If set to \c VARIANT_TRUE, the event gets fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessReturnKey, Raise_OpenLink
	virtual HRESULT STDMETHODCALLTYPE put_ProcessReturnKey(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. If set to
	/// \c VARIANT_TRUE, the control accepts OLE drag'n'drop actions; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RightToLeft, LblCtlsLibU::RightToLeftConstants
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RightToLeft, LblCtlsLibU::RightToLeftConstants
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowToolTips property</em>
	///
	/// Retrieves whether the control displays any tooltips. If set to \c VARIANT_TRUE, tooltips are
	/// displayed for links, with the link's target as default text. This text can be customized by
	/// handling the \c LinkGetInfoTipText event. If set to \c VARIANT_FALSE, no tooltips are shown.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowToolTips, get_hWndToolTip, Raise_LinkGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE get_ShowToolTips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowToolTips property</em>
	///
	/// Sets whether the control displays any tooltips. If set to \c VARIANT_TRUE, tooltips are
	/// displayed for links, with the link's target as default text. This text can be customized by
	/// handling the \c LinkGetInfoTipText event. If set to \c VARIANT_FALSE, no tooltips are shown.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowToolTips, put_hWndToolTip, Raise_LinkGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE put_ShowToolTips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the control's content.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa put_Text, get_ForeColor, get_Font, get_HAlignment, get_UseMnemonic, GetIdealSize,
	///     Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the control's content.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the control's default property.
	///
	/// \sa get_Text, put_ForeColor, put_Font, put_HAlignment, put_UseMnemonic, GetIdealSize,
	///     Raise_TextChanged
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c UseMnemonic property</em>
	///
	/// Retrieves whether the control handles the last ampersand in the control text as accelerator key and
	/// hides any other ampersands, except if they are escaped by doubling them, i.e. the text contains
	/// &quot;&amp;&amp;&quot; instead of &quot;&amp;&quot;. If set to \c VARIANT_TRUE, ampersands are
	/// handled as marking the accelerator key; otherwise they are handled as normal characters.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_UseMnemonic, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_UseMnemonic(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseMnemonic property</em>
	///
	/// Sets whether the control handles the last ampersand in the control text as accelerator key and
	/// hides any other ampersands, except if they are escaped by doubling them, i.e. the text contains
	/// &quot;&amp;&amp;&quot; instead of &quot;&amp;&quot;. If set to \c VARIANT_TRUE, ampersands are
	/// handled as marking the accelerator key; otherwise they are handled as normal characters.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_UseMnemonic, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_UseMnemonic(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseSystemFont, get_Font, get_UseVisualStyle
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font, put_UseVisualStyle
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseVisualStyle property</em>
	///
	/// Retrieves whether the control is drawn using the current Visual Style. If set to \c VARIANT_TRUE,
	/// colors and the font are defined by the current Visual Style, overriding the appropriate properties;
	/// otherwise the colors and font are defined by the properties.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_UseVisualStyle, get_BackColor, get_ForeColor, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseVisualStyle(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseVisualStyle property</em>
	///
	/// Sets whether the control is drawn using the current Visual Style. If set to \c VARIANT_TRUE,
	/// colors and the font are defined by the current Visual Style, overriding the appropriate properties;
	/// otherwise the colors and font are defined by the properties.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_UseVisualStyle, put_BackColor, put_ForeColor, put_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseVisualStyle(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Retrieves the control's optimal size</em>
	///
	/// Calculates the rectangle that best fits the control's text.
	///
	/// \param[in] maximumWidth The control's maximum width in pixels. If set to -1, the control's current
	///            width is used. If set to 0, it is assumed that all text should be put into a single line.
	/// \param[out] pIdealWidth The control's optimal width in pixels.
	/// \param[out] pIdealHeight The control's optimal height in pixels.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text
	virtual HRESULT STDMETHODCALLTYPE GetIdealSize(long maximumWidth = 0, OLE_XSIZE_PIXELS* pIdealWidth = NULL, OLE_YSIZE_PIXELS* pIdealHeight = NULL);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in,out] pHitTestDetails Receives a value specifying the exact part of the control the
	///                specified point lies in. Any of the values defined by the \c HitTestConstants
	///                enumeration is valid.
	/// \param[out] ppHitLink Receives the "hit" link's \c ILink implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LblCtlsLibU::HitTestConstants, HitTest
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, ILink** ppHitLink);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	/// \param[in] languageID The locale identifier identifying the language in which name should be
	///            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties, StringProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KILLFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnSetFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646282.aspx">WM_KILLFOCUS</a>
	LRESULT OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to relay the message to the tooltip control.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, ToolTipStatus::RelayToToolTip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to relay the message to the tooltip control.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, ToolTipStatus::RelayToToolTip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnSetFocus, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NCHITTEST handler</em>
	///
	/// Will be called if the Windows window engine needs to know what the specified point belongs to.
	/// We use this handler to support mouse events and for the \c MousePointer and \c MouseIcon properties.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms645618.aspx">WM_NCHITTEST</a>
	LRESULT OnNCHitTest(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_NCMOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// non-client area.
	/// We use this handler to relay the message to the tooltip control.
	///
	/// \sa ToolTipStatus::RelayToToolTip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645627.aspx">WM_NCMOUSEMOVE</a>
	LRESULT OnNCMouseMove(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to relay the message to the tooltip control.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, ToolTipStatus::RelayToToolTip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTEXT handler</em>
	///
	/// Will be called if the control's text shall be changed.
	/// We use this handler to support data-binding for the \c Text property and to raise the \c TextChanged
	/// event.
	///
	/// \sa get_Text, Raise_TextChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632644.aspx">WM_SETTEXT</a>
	LRESULT OnSetText(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SIZE handler</em>
	///
	/// Will be called if the control's size was changed.
	/// We use this handler to update the tooltip control.
	///
	/// \sa OnWindowPosChanged, ToolTipStatus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632646.aspx">WM_SIZE</a>
	LRESULT OnSize(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa OnSize, Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CTLCOLORSTATIC handler</em>
	///
	/// Will be called if the control asks its parent window to configure the specified device context for
	/// drawing the control, i. e. to setup the colors and brushes.
	/// We use this handler for proper theming support and for the \c BackColor and \c ForeColor properties.
	///
	/// \sa get_BackColor, get_ForeColor,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms651165.aspx">WM_CTLCOLORSTATIC</a>
	LRESULT OnReflectedCtlColorStatic(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the control's parent window receives a notification from the control.
	/// We use this handler for the custom draw support.
	///
	/// \sa OnCustomDrawNotification, OnCustomTextNotification,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775583.aspx">WM_NOTIFY</a>
	LRESULT OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c LM_SETITEM handler</em>
	///
	/// Will be called if the control is requested to update some or all of a link's attributes.
	/// We use this handler to detect changes of the caret link.
	///
	/// \sa Raise_CaretChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760724.aspx">LM_SETITEM</a>
	LRESULT OnSetItem(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c NM_CLICK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked into the control's
	/// client area using the left mouse button.
	/// We use this handler to raise the \c OpenLink event.
	///
	/// \sa OnLButtonDown, Raise_OpenLink,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760714.aspx">NM_CLICK (syslink)</a>
	LRESULT OnClickNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_CUSTOMTEXT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's content is being drawn.
	/// The client application use this notification to customize the text and its appearance. The difference
	/// to the \c NM_CUSTOMDRAW notification is that \c NM_CUSTOMTEXT always is raised on a per-text-block
	/// basis and also provides the bounding rectangle of each single text block.\n
	/// We use this handler to raise the \c CustomizeTextDrawing event.
	///
	/// \sa OnCustomDrawNotification, Raise_CustomizeTextDrawing,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775530.aspx">NM_CUSTOMTEXT</a>
	LRESULT OnCustomTextNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_RETURN handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user pressed the \c [RETURN] key
	/// while the control has the focus.
	/// We use this handler to raise the \c OpenLink event.
	///
	/// \sa OnKeyDown, Raise_OpenLink,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775562.aspx">NM_RETURN</a>
	LRESULT OnReturnNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called if the control is queried by the tooltip control for the tooltip text to display.
	/// We use this handler to provide the appropriate tooltip text.
	///
	/// \sa OnToolTipGetDispInfoNotificationW,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760269.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called if the control is queried by the tooltip control for the tooltip text to display.
	/// We use this handler to provide the appropriate tooltip text.
	///
	/// \sa OnToolTipGetDispInfoNotificationA,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760269.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>\c NM_CUSTOMDRAW handler</em>
	///
	/// Will be called by the \c OnReflectedNotify method if a custom draw notification was received for
	/// the control.
	/// We use this handler to raise the \c CustomDraw event.
	///
	/// \sa OnReflectedNotify, Raise_CustomDraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775528.aspx">NM_CUSTOMDRAW</a>
	LRESULT OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c CaretChanged event</em>
	///
	/// \param[in] previousCaretLink The previous caret link.
	/// \param[in] newCaretLink The new caret link.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_CaretChanged, LblCtlsLibU::_ISysLinkEvents::CaretChanged,
	///     Raise_OpenLink, get_CaretLink
	inline HRESULT Raise_CaretChanged(int previousCaretLink, int newCaretLink);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_Click, LblCtlsLibU::_ISysLinkEvents::Click, Raise_OpenLink,
	///     Raise_DblClick, Raise_MClick, Raise_RClick, Raise_XClick
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_ContextMenu, LblCtlsLibU::_ISysLinkEvents::ContextMenu, Raise_RClick
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CustomDraw event</em>
	///
	/// \param[in] pLink The link that the notification refers to. Will be \c NULL, if the notification
	///            refers to the whole control content or to text that is not a link.
	/// \param[in] nonLinkPartIndex A zero-based index that identifies to which plain-text part of the
	///            content the event refers to. If the notification refers to the whole control content or to
	///            a link, this parameter will be \c -1.
	/// \param[in,out] pTextColor An \c OLE_COLOR value specifying the color to draw the text in. The client
	///                may change this value.
	/// \param[in,out] pTextBackColor An \c OLE_COLOR value specifying the color to fill the text background
	///                with. The client may change this value.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Any of the values
	///            defined by the \c CustomDrawStageConstants enumeration is valid.
	/// \param[in] linkState The link's current state (focused, disabled etc.). Most of the values
	///            defined by the \c CustomDrawLinkStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the area
	///            that needs to be drawn. If the notification does not refer to the whole control, the
	///            rectangle will be empty.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.\n
	///          For the control content "This is <a>a link</a> and <a>another link</a> and some text" the
	///          \c link and \c nonLinkPartIndex parameters will be as follows if the event is raised to
	///          customize drawing of single parts of the content:
	///          <pre>
	///            Event No.  Part           link             nonLinkPartIndex
	///            1          This is        NULL              0
	///            2          a link         Object (Index 0) -1
	///            3          and            NULL              1
	///            4          another link   Object (Index 1) -1
	///            5          and some text  NULL              2
	///          </pre>
	///
	/// \sa Proxy_ISysLinkEvents::Fire_CustomDraw, LblCtlsLibU::_ISysLinkEvents::CustomDraw,
	///     Raise_CustomizeTextDrawing, get_ForeColor, LblCtlsLibU::RECTANGLE,
	///     LblCtlsLibU::CustomDrawStageConstants, LblCtlsLibU::CustomDrawLinkStateConstants,
	///     LblCtlsLibU::CustomDrawReturnValuesConstants,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775528.aspx">NM_CUSTOMDRAW</a>
	inline HRESULT Raise_CustomDraw(ILink* pLink, LONG nonLinkPartIndex, OLE_COLOR* pTextColor, OLE_COLOR* pTextBackColor, CustomDrawStageConstants drawStage, CustomDrawLinkStateConstants linkState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing);
	/// \brief <em>Raises the \c CustomizeTextDrawing event</em>
	///
	/// \param[in] isLink The text specified by \c pText is a link text.
	/// \param[in] maxTextLength The maximum number of characters the block's text may consist of.
	/// \param[in,out] pText The text of the block currently being drawn.
	/// \param[in] hDC The handle of the device context in which all drawing should take place.
	/// \param[in,out] pDrawingRectangle The bounding rectangle of the text block that is being drawn.
	/// \param[in] drawTextFlags A combination of \c DT_* flags specifying how the text will be drawn. Any
	///            combination of the \c DT_* flags that can be used with the \c DrawText API function is
	///            valid.
	/// \param[in,out] pDrawText If set to \c VARIANT_TRUE, the text block will be drawn; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.\n
	///          Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_CustomizeTextDrawing,
	///     LblCtlsLibU::_ISysLinkEvents::CustomizeTextDrawing, Raise_CustomDraw, LblCtlsLibU::RECTANGLE,
	///     get_DisabledEvents,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775530.aspx">NM_CUSTOMTEXT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd162498.aspx">DrawText</a>
	inline HRESULT Raise_CustomizeTextDrawing(VARIANT_BOOL isLink, LONG maxTextLength, BSTR* pText, LONG hDC, RECTANGLE* pDrawingRectangle, LONG drawTextFlags, VARIANT_BOOL* pDrawText);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_DblClick, LblCtlsLibU::_ISysLinkEvents::DblClick, Raise_Click,
	///     Raise_MDblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_DestroyedControlWindow,
	///     LblCtlsLibU::_ISysLinkEvents::DestroyedControlWindow, Raise_RecreatedControlWindow, get_hWnd
	inline HRESULT Raise_DestroyedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_KeyDown, LblCtlsLibU::_ISysLinkEvents::KeyDown, Raise_KeyUp,
	///     Raise_KeyPress
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_KeyPress, LblCtlsLibU::_ISysLinkEvents::KeyPress, Raise_KeyDown,
	///     Raise_KeyUp
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_KeyUp, LblCtlsLibU::_ISysLinkEvents::KeyUp, Raise_KeyDown,
	///     Raise_KeyPress
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c LinkGetInfoTipText event</em>
	///
	/// \param[in] pLink The link whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The link's info tip text.
	/// \param[in,out] pAbortToolTip If \c VARIANT_TRUE, the caller should NOT show the tooltip;
	///                otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_LinkGetInfoTipText,
	///     LblCtlsLibU::_ISysLinkEvents::LinkGetInfoTipText, put_ShowToolTips
	inline HRESULT Raise_LinkGetInfoTipText(ILink* pLink, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip);
	/// \brief <em>Raises the \c LinkMouseEnter event</em>
	///
	/// \param[in] pLink The link that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_LinkMouseEnter, LblCtlsLibU::_ISysLinkEvents::LinkMouseEnter,
	///     Raise_LinkMouseLeave, Raise_MouseMove, LblCtlsLibU::HitTestConstants
	inline HRESULT Raise_LinkMouseEnter(ILink* pLink, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c LinkMouseLeave event</em>
	///
	/// \param[in] pLink The link that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_LinkMouseLeave, LblCtlsLibU::_ISysLinkEvents::LinkMouseLeave,
	///     Raise_LinkMouseEnter, Raise_MouseMove, LblCtlsLibU::HitTestConstants
	inline HRESULT Raise_LinkMouseLeave(ILink* pLink, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MClick, LblCtlsLibU::_ISysLinkEvents::MClick, Raise_MDblClick,
	///     Raise_Click, Raise_RClick, Raise_XClick
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MDblClick, LblCtlsLibU::_ISysLinkEvents::MDblClick, Raise_MClick,
	///     Raise_DblClick, Raise_RDblClick, Raise_XDblClick, get_DetectDoubleClicks
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MouseDown, LblCtlsLibU::_ISysLinkEvents::MouseDown, Raise_MouseUp,
	///     Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MouseEnter, LblCtlsLibU::_ISysLinkEvents::MouseEnter,
	///     Raise_MouseLeave, Raise_LinkMouseEnter, Raise_MouseHover, Raise_MouseMove
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MouseHover, LblCtlsLibU::_ISysLinkEvents::MouseHover,
	///     Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MouseLeave, LblCtlsLibU::_ISysLinkEvents::MouseLeave,
	///     Raise_MouseEnter, Raise_LinkMouseLeave, Raise_MouseHover, Raise_MouseMove
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MouseMove, LblCtlsLibU::_ISysLinkEvents::MouseMove,
	///     Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_MouseUp, LblCtlsLibU::_ISysLinkEvents::MouseUp, Raise_MouseDown,
	///     Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_OLEDragDrop, LblCtlsLibU::_ISysLinkEvents::OLEDragDrop,
	///     Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///     get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_OLEDragEnter, LblCtlsLibU::_ISysLinkEvents::OLEDragEnter,
	///     Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_OLEDragLeave, LblCtlsLibU::_ISysLinkEvents::OLEDragLeave,
	///     Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_OLEDragMouseMove, LblCtlsLibU::_ISysLinkEvents::OLEDragMouseMove,
	///     Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OpenLink event</em>
	///
	/// \param[in] pLink The link that should be opened.
	/// \param[in] causedBy Specifies how the event has been triggered. Any of the values defined by the
	///            \c EventCausedByConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_OpenLink, LblCtlsLibU::_ISysLinkEvents::OpenLink, Raise_Click,
	///     LblCtlsLibU::EventCausedByConstants, get_ProcessReturnKey, Link::put_Visited
	inline HRESULT Raise_OpenLink(ILink* pLink, EventCausedByConstants causedBy);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_RClick, LblCtlsLibU::_ISysLinkEvents::RClick, Raise_ContextMenu,
	///     Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_RDblClick, LblCtlsLibU::_ISysLinkEvents::RDblClick, Raise_RClick,
	///     Raise_DblClick, Raise_MDblClick, Raise_XDblClick, get_DetectDoubleClicks
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_RecreatedControlWindow,
	///     LblCtlsLibU::_ISysLinkEvents::RecreatedControlWindow, Raise_DestroyedControlWindow, get_hWnd
	inline HRESULT Raise_RecreatedControlWindow(HWND hWnd);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_ResizedControlWindow,
	///     LblCtlsLibU::_ISysLinkEvents::ResizedControlWindow
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c TextChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_TextChanged, LblCtlsLibU::_ISysLinkEvents::TextChanged
	inline HRESULT Raise_TextChanged(void);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_XClick, LblCtlsLibU::_ISysLinkEvents::XClick,
	///     Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \sa Proxy_ISysLinkEvents::Fire_XDblClick, LblCtlsLibU::_ISysLinkEvents::XDblClick,
	///     Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick, get_DetectDoubleClicks
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnMnemonic, Properties::hAcceleratorTable,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	/// \brief <em>Implementation of \c IOleControl::OnMnemonic</em>
	///
	/// Will be called if the user has pressed a keystroke that matches one of the entries in the control's
	/// accelerator table.
	///
	/// \param[in] pMessage The \c MSG structure describing the keystroke.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetControlInfo, Properties::hAcceleratorTable,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680699.aspx">IOleControl::OnMnemonic</a>
	virtual HRESULT STDMETHODCALLTYPE OnMnemonic(LPMSG /*pMessage*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	/// \brief <em>Increments the \c silentCaretChanges flag</em>
	///
	/// \sa LeaveSilentCaretChangeSection, Flags::silentCaretChanges, Raise_CaretChanged
	void EnterSilentCaretChangeSection(void);
	/// \brief <em>Decrements the \c silentCaretChanges flag</em>
	///
	/// \sa EnterSilentCaretChangeSection, Flags::silentCaretChanges, Raise_CaretChanged
	void LeaveSilentCaretChangeSection(void);

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c CB_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \sa LblCtlsLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the link that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of SLHT_* flags, that holds further details about the control's
	///             part below the specified point.
	///
	/// \return The "hit" link's index.
	///
	/// \sa HitTest
	int HitTest(LONG x, LONG y, UINT* pFlags);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);


	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<SysLink> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(SysLink* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<SysLink> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<SysLink> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(SysLink* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<SysLink> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
				ATLASSERT(watching == 0);
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutomaticallyMarkLinksAsVisited property's setting</em>
		///
		/// \sa get_AutomaticallyMarkLinksAsVisited, put_AutomaticallyMarkLinksAsVisited
		UINT automaticallyMarkLinksAsVisited : 1;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BackStyle property's setting</em>
		///
		/// \sa get_BackStyle, put_BackStyle
		BackStyleConstants backStyle;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c DetectDoubleClicks property's setting</em>
		///
		/// \sa get_DetectDoubleClicks, put_DetectDoubleClicks
		UINT detectDoubleClicks : 1;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the \c HAlignment property's setting</em>
		///
		/// \sa get_HAlignment, put_HAlignment
		HAlignmentConstants hAlignment;
		/// \brief <em>Holds the \c HotTracking property's setting</em>
		///
		/// \sa get_HotTracking, put_HotTracking
		UINT hotTracking : 1;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c ProcessReturnKey property's setting</em>
		///
		/// \sa get_ProcessReturnKey, put_ProcessReturnKey
		UINT processReturnKey : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		UINT registerForOLEDragDrop : 1;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ShowToolTips property's setting</em>
		///
		/// \sa get_ShowToolTips, put_ShowToolTips
		UINT showToolTips : 1;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the control's accelerator character</em>
		///
		/// \sa OnSetText, hAcceleratorTable, GetControlInfo, OnMnemonic
		TCHAR accelerator;
		/// \brief <em>Holds the control's accelerator table</em>
		///
		/// \sa OnSetText, accelerator, GetControlInfo, OnMnemonic
		HACCEL hAcceleratorTable;
		/// \brief <em>If \c TRUE, \c Create will set \c text to the control's name</em>
		///
		/// \sa Create, text
		UINT resetTextToName : 1;
		/// \brief <em>Holds the \c Text property's setting</em>
		///
		/// \sa get_Text, put_Text, resetTextToName
		CComBSTR text;
		/// \brief <em>Holds the \c UseMnemonic property's setting</em>
		///
		/// \sa get_UseMnemonic, put_UseMnemonic
		UINT useMnemonic : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c UseVisualStyle property's setting</em>
		///
		/// \sa get_UseVisualStyle, put_UseVisualStyle
		UINT useVisualStyle : 1;

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			if(hAcceleratorTable) {
				DestroyAcceleratorTable(hAcceleratorTable);
				hAcceleratorTable = NULL;
			}
			font.Release();
			mouseIcon.Release();
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			appearance = a2D;
			automaticallyMarkLinksAsVisited = TRUE;
			backColor = 0x80000000 | COLOR_BTNFACE;
			backStyle = bstOpaque;
			borderStyle = bsNone;
			detectDoubleClicks = TRUE;
			disabledEvents = static_cast<DisabledEventsConstants>(deClickEvents | deMouseEvents | deKeyboardEvents | deCustomDraw | deTextChangedEvents | deCustomizeTextDrawing);
			dontRedraw = FALSE;
			enabled = TRUE;
			foreColor = 0x80000000 | COLOR_BTNTEXT;
			hAlignment = halLeft;
			hotTracking = TRUE;
			hoverTime = -1;
			mousePointer = mpDefault;
			processContextMenuKeys = TRUE;
			processReturnKey = TRUE;
			registerForOLEDragDrop = FALSE;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			showToolTips = TRUE;
			supportOLEDragImages = TRUE;
			text = L"<a>SysLink</a>";
			accelerator = L'\0';
			hAcceleratorTable = NULL;
			resetTextToName = TRUE;
			useMnemonic = TRUE;
			useSystemFont = TRUE;
			useVisualStyle = FALSE;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the control has been activated by mouse and needs to be UI-activated by \c OnSetFocus</em>
		///
		/// ATL always UI-activates the control in \c OnMouseActivate. If the control is activated by mouse,
		/// \c WM_SETFOCUS is sent after \c WM_MOUSEACTIVATE, but Visual Basic 6 won't raise the \c Validate
		/// event if the control already is UI-activated when it receives the focus. Therefore we need to delay
		/// UI-activation.
		///
		/// \sa OnMouseActivate, OnSetFocus, OnKillFocus
		UINT uiActivationPending : 1;
		/// \brief <em>If \c TRUE, \c RecreateControlWindow won't recreate the control window</em>
		///
		/// \sa RecreateControlWindow, Load
		UINT dontRecreate : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;
		/// \brief <em>If not 0, we won't raise \c CaretChanged event</em>
		///
		/// \sa EnterSilentCaretChangeSection, LeaveSilentCaretChangeSection, Raise_CaretChanged
		int silentCaretChanges;

		Flags()
		{
			uiActivationPending = FALSE;
			dontRecreate = FALSE;
			usingThemes = FALSE;
			silentCaretChanges = 0;
		}
	} flags;

	/// \brief <em>Holds the index of the current caret link</em>
	int currentCaretLink;
	/// \brief <em>Holds the index of the link below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deMouseEvents being set.
	int linkUnderMouse;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;
		/// \brief <em>Holds the mouse button that may cause a double-click on the next \c WM_*BUTTONDOWN message</em>
		///
		/// \sa StoreDoubleClickCandidate, IsDoubleClickCandidate, RemoveAllDoubleClickCandidates,
		///     Raise_MouseDown, Raise_MouseUp
		SHORT doubleClickCandidate;
		/// \brief <em>Holds the timestamp until which the next \c WM_*BUTTONDOWN message has to be received in order to trigger a double-click</em>
		///
		/// \sa StoreDoubleClickCandidate, IsDoubleClickCandidate, RemoveAllDoubleClickCandidates,
		///     Raise_MouseDown, Raise_MouseUp
		LONG doubleClickTimeOut;
		/// \brief <em>Holds the position of the click message that has been stored as a candidate for a double-click</em>
		///
		/// \sa StoreDoubleClickCandidate, IsDoubleClickCandidate, RemoveAllDoubleClickCandidates,
		///     Raise_MouseDown, Raise_MouseUp
		POINT clickPosition;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;
		/// \brief <em>Holds the index of the last clicked link</em>
		///
		/// Holds the index of the last clicked link. We use this to ensure that the \c *DblClick events
		/// are not raised accidently.
		///
		/// \attention This member is not reliable with \c deClickEvents being set.
		///
		/// \sa Raise_Click, Raise_DblClick, Raise_MClick, Raise_MDblClick, Raise_RClick, Raise_RDblClick,
		///     Raise_XClick, Raise_XDblClick
		int lastClickedLink;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedLink = -1;
			doubleClickCandidate = 0;
			doubleClickTimeOut = 0;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
			lastClickedLink = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedLink = -1;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Stores a mouse button as double-click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa doubleClickCandidate, IsDoubleClickCandidate, RemoveAllDoubleClickCandidates
		void StoreDoubleClickCandidate(SHORT button)
		{
			doubleClickCandidate = button;
			doubleClickTimeOut = GetMessageTime() + GetDoubleClickTime();
			DWORD position = GetMessagePos();
			clickPosition.x = GET_X_LPARAM(position);
			clickPosition.y = GET_Y_LPARAM(position);
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Retrieves whether a mouse button is a double-click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a double-click candidate and the current message also
		///         matches the other double-click conditions; otherwise \c FALSE.
		///
		/// \sa doubleClickCandidate, StoreDoubleClickCandidate, RemoveAllDoubleClickCandidates
		BOOL IsDoubleClickCandidate(SHORT button)
		{
			if(IsClickCandidate(button) && doubleClickCandidate == button && GetMessageTime() <= doubleClickTimeOut) {
				DWORD position = GetMessagePos();
				if(abs(GET_X_LPARAM(position) - clickPosition.x) < (GetSystemMetrics(SM_CXDOUBLECLK) >> 1) && abs(GET_Y_LPARAM(position) - clickPosition.y) < (GetSystemMetrics(SM_CYDOUBLECLK) >> 1)) {
					return TRUE;
				}
			}
			return FALSE;
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa doubleClickCandidate, StoreDoubleClickCandidate, IsDoubleClickCandidate
		void RemoveAllDoubleClickCandidates(void)
		{
			doubleClickCandidate = 0;
			doubleClickTimeOut = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pDropTargetHelper = NULL;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
			}
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			drop_pDataObject = NULL;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			//
		}
	} dragDropStatus;

	/// \brief <em>Holds data and flags related to tooltip support</em>
	struct ToolTipStatus
	{
		/// \brief <em>The tooltip control to use</em>
		HWND hWndToolTip;
		/// \brief <em>The link for which the tooltip is currently displayed</em>
		int toolTipLink;
		/// \brief <em>Holds the current tooltip text</em>
		LPSTR pCurrentToolTipA;
		/// \brief <em>Holds the current tooltip text</em>
		LPWSTR pCurrentToolTipW;

		ToolTipStatus()
		{
			hWndToolTip = NULL;
			toolTipLink = -1;
			pCurrentToolTipA = NULL;
			pCurrentToolTipW = NULL;
		}

		~ToolTipStatus()
		{
			if(::IsWindow(hWndToolTip)) {
				::DestroyWindow(hWndToolTip);
			}
			
			typedef BOOL WINAPI Str_SetPtrAFn(__deref_inout_opt LPSTR*, __in_opt LPCSTR);
			Str_SetPtrAFn* pfnStr_SetPtrA = NULL;
			HMODULE hComctl32DLL = LoadLibrary(TEXT("comctl32.dll"));
			if(hComctl32DLL) {
				pfnStr_SetPtrA = reinterpret_cast<Str_SetPtrAFn*>(GetProcAddress(hComctl32DLL, MAKEINTRESOURCEA(234)));
				FreeLibrary(hComctl32DLL);
			}
			if(pfnStr_SetPtrA) {
				pfnStr_SetPtrA(&pCurrentToolTipA, NULL);
			}
			Str_SetPtrW(&pCurrentToolTipW, NULL);
		}

		/// \brief <em>Relays a message to the tooltip control</em>
		///
		/// \param[in] hWnd The original recipient of the message to relay.
		/// \param[in] message The message to relay.
		/// \param[in] wParam The first parameter of the message to relay.
		/// \param[in] lParam The second parameter of the message to relay.
		void RelayToToolTip(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
		{
			if(hWndToolTip) {
				MSG msg;
				msg.lParam = lParam;
				msg.wParam = wParam;
				msg.message = message;
				msg.hwnd = hWnd;
				SendMessage(hWndToolTip, TTM_RELAYEVENT, 0, reinterpret_cast<LPARAM>(&msg));
			}
		}

		/// \brief <em>Invalidates the \c toolTipLink member</em>
		void InvalidateToolTipLink(void)
		{
			toolTipLink = -1;
			if(pCurrentToolTipA && pCurrentToolTipA != LPSTR_TEXTCALLBACKA) {
				pCurrentToolTipA[0] = 0;
			}
			if(pCurrentToolTipW && pCurrentToolTipW != LPSTR_TEXTCALLBACKW) {
				pCurrentToolTipW[0] = 0;
			}
		}
	} toolTipStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
	} timers;

	/// \brief <em>The zero-based index of the control content part currently being custom-drawn</em>
	///
	/// \sa OnCustomDrawNotification, Raise_CustomDraw
	LONG customDrawNonLinkPartIndex;

	//////////////////////////////////////////////////////////////////////
	/// \name Version information
	///
	//@{
	/// \brief <em>Retrieves whether we're using at least version 6.10 of comctl32.dll</em>
	///
	/// \return \c TRUE if we're using comctl32.dll version 6.10 or higher; otherwise \c FALSE.
	BOOL IsComctl32Version610OrNewer(void);
	//@}
	//////////////////////////////////////////////////////////////////////

private:
};     // SysLink

OBJECT_ENTRY_AUTO(__uuidof(SysLink), SysLink)