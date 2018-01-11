//////////////////////////////////////////////////////////////////////
/// \class Link
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing hyper link</em>
///
/// This class is a wrapper around a hyper link that really exists in the control.
///
/// \sa LblCtlsLibU::ILink, Links, SysLink
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#include "LblCtlsU.h"
#include "_ILinkEvents_CP.h"
#include "helpers.h"
#include "SysLink.h"


class ATL_NO_VTABLE Link : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<Link, &CLSID_Link>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<Link>,
    public Proxy_ILinkEvents<Link>, 
   	public IDispatchImpl<ILink, &IID_ILink, &LIBID_LblCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
{
	friend class SysLink;
	friend class Links;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_LINK)

		BEGIN_COM_MAP(Link)
			COM_INTERFACE_ENTRY(ILink)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(Link)
			CONNECTION_POINT_ENTRY(__uuidof(_ILinkEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
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
	/// \name Implementation of ILink
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Caret property</em>
	///
	/// Retrieves whether the link is the control's caret link, i. e. it has the focus. If it is the
	/// caret link, this property is set to \c VARIANT_TRUE; otherwise it's set to \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Enabled, SysLink::get_CaretLink
	virtual HRESULT STDMETHODCALLTYPE get_Caret(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c DrawAsNormalText property</em>
	///
	/// Retrieves whether the link is drawn like a link or like normal text. If this property is set to
	/// \c VARIANT_TRUE, the link is drawn as normal text, but still underlined; otherwise it is drawn like
	/// a link, i.e. in a different color.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_DrawAsNormalText, SysLink::get_ForeColor, get_Enabled, get_Hot, get_Visited
	virtual HRESULT STDMETHODCALLTYPE get_DrawAsNormalText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DrawAsNormalText property</em>
	///
	/// Sets whether the link is drawn like a link or like normal text. If this property is set to
	/// \c VARIANT_TRUE, the link is drawn as normal text, but still underlined; otherwise it is drawn like
	/// a link, i.e. in a different color.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_DrawAsNormalText, SysLink::put_ForeColor, put_Enabled, put_Hot, put_Visited
	virtual HRESULT STDMETHODCALLTYPE put_DrawAsNormalText(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the link is enabled, i. e. whether it can be clicked. If this property is set to
	/// \c VARIANT_TRUE, the link is enabled; otherwise it is disabled.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, get_Hot, get_Visited, SysLink::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Sets whether the link is enabled, i. e. whether it can be clicked. If this property is set to
	/// \c VARIANT_TRUE, the link is enabled; otherwise it is disabled.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, put_Hot, put_Visited, SysLink::put_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Hot property</em>
	///
	/// Retrieves whether the link is displayed as hot-tracked, i. e. whether it is displayed as if the
	/// mouse cursor is located over it. If this property is set to \c VARIANT_TRUE, the link is hot-tracked;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Hot, get_Enabled, get_Visited
	virtual HRESULT STDMETHODCALLTYPE get_Hot(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Hot property</em>
	///
	/// Sets whether the link is displayed as hot-tracked, i. e. whether it is displayed as if the
	/// mouse cursor is located over it. If this property is set to \c VARIANT_TRUE, the link is hot-tracked;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Hot, put_Enabled, put_Visited
	virtual HRESULT STDMETHODCALLTYPE put_Hot(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves the link's ID.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ID, get_URL, get_Text, get_Index, LblCtlsLibU::LinkIdentifierTypeConstants
	virtual HRESULT STDMETHODCALLTYPE get_ID(BSTR* pValue);
	/// \brief <em>Sets the \c ID property</em>
	///
	/// Sets the link's ID.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ID, put_URL, put_Text, LblCtlsLibU::LinkIdentifierTypeConstants
	virtual HRESULT STDMETHODCALLTYPE put_ID(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this link.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ID, LblCtlsLibU::LinkIdentifierTypeConstants
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the link's text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, get_ID, get_URL
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the link's text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, put_ID, put_URL
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c URL property</em>
	///
	/// Retrieves the link's URL.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_URL, get_ID, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_URL(BSTR* pValue);
	/// \brief <em>Sets the \c URL property</em>
	///
	/// Sets the link's URL.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_URL, put_ID, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_URL(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c Visited property</em>
	///
	/// Retrieves whether the link is drawn like a link that already has been visited. If this property is
	/// set to \c VARIANT_TRUE, the link is drawn as visited; otherwise it is drawn like a new link.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Visited, get_Hot, get_Enabled, SysLink::Raise_OpenLink
	virtual HRESULT STDMETHODCALLTYPE get_Visited(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Visited property</em>
	///
	/// Sets whether the link is drawn like a link that already has been visited. If this property is
	/// set to \c VARIANT_TRUE, the link is drawn as visited; otherwise it is drawn like a new link.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Visited, put_Hot, put_Enabled, SysLink::Raise_OpenLink
	virtual HRESULT STDMETHODCALLTYPE put_Visited(VARIANT_BOOL newValue);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given link</em>
	///
	/// Attaches this object to a given link, so that the link's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] linkIndex The link to attach to.
	///
	/// \sa Detach
	void Attach(int linkIndex);
	/// \brief <em>Detaches this object from a link</em>
	///
	/// Detaches this object from the link it currently wraps, so that it doesn't wrap any link anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets the owner of this link</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerSysLink
	void SetOwner(__in_opt SysLink* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c SysLink object that owns this link</em>
		///
		/// \sa SetOwner
		SysLink* pOwnerSysLink;
		/// \brief <em>The index of the link wrapped by this object</em>
		int linkIndex;

		Properties()
		{
			pOwnerSysLink = NULL;
			linkIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning sys link control's window handle</em>
		///
		/// \return The window handle of the sys link control that contains this link.
		///
		/// \sa pOwnerSysLink
		HWND GetSysLinkHWnd(void);
	} properties;
};     // Link

OBJECT_ENTRY_AUTO(__uuidof(Link), Link)