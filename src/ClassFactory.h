//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \if UNICODE
///   \sa SysLink, WindowedLabel, WindowlessLabel
/// \else
///   \sa WindowedLabel, WindowlessLabel
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#ifdef UNICODE
	#include "Link.h"
	#include "Links.h"
#endif
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	#ifdef UNICODE
		/// \brief <em>Creates a \c Link object</em>
		///
		/// Creates a \c Link object that represents a given hyper link and returns its \c ILink implementation
		/// as a smart pointer.
		///
		/// \param[in] linkIndex The index of the link for which to create the object.
		/// \param[in] pSysLink The \c SysLink object the \c Link object will work on.
		/// \param[in] validateLinkIndex If \c TRUE, the method will check whether the \c linkIndex parameter
		///            identifies an existing link; otherwise not.
		///
		/// \return A smart pointer to the created object's \c ILink implementation.
		static inline CComPtr<ILink> InitLink(int linkIndex, SysLink* pSysLink, BOOL validateLinkIndex = TRUE)
		{
			CComPtr<ILink> pLink = NULL;
			InitLink(linkIndex, pSysLink, IID_ILink, reinterpret_cast<IUnknown**>(&pLink), validateLinkIndex);
			return pLink;
		};

		/// \brief <em>Creates a \c Link object</em>
		///
		/// \overload
		///
		/// Creates a \c Link object that represents a given hyper link and returns its implementation of a
		/// given interface as a raw pointer.
		///
		/// \param[in] linkIndex The index of the link for which to create the object.
		/// \param[in] pSysLink The \c SysLink object the \c Link object will work on.
		/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
		///            be returned.
		/// \param[out] ppLink Receives the object's implementation of the interface identified by
		///             \c requiredInterface.
		/// \param[in] validateLinkIndex If \c TRUE, the method will check whether the \c linkIndex parameter
		///            identifies an existing link; otherwise not.
		static inline void InitLink(int linkIndex, SysLink* pSysLink, REFIID requiredInterface, IUnknown** ppLink, BOOL validateLinkIndex = TRUE)
		{
			ATLASSERT_POINTER(ppLink, IUnknown*);
			ATLASSUME(ppLink);

			*ppLink = NULL;
			if(validateLinkIndex && !IsValidLinkIndex(linkIndex, pSysLink)) {
				// there's nothing useful we could return
				return;
			}

			// create a Link object and initialize it with the given parameters
			CComObject<Link>* pLinkObj = NULL;
			CComObject<Link>::CreateInstance(&pLinkObj);
			pLinkObj->AddRef();
			pLinkObj->SetOwner(pSysLink);
			pLinkObj->Attach(linkIndex);
			pLinkObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppLink));
			pLinkObj->Release();
		};

		/// \brief <em>Creates a \c Links object</em>
		///
		/// Creates a \c Links object that represents a collection of hyper links and returns its \c ILinks
		/// implementation as a smart pointer.
		///
		/// \param[in] pSysLink The \c SysLink object the \c Links object will work on.
		///
		/// \return A smart pointer to the created object's \c ILinks implementation.
		static inline CComPtr<ILinks> InitLinks(SysLink* pSysLink)
		{
			CComPtr<ILinks> pLinks = NULL;
			InitLinks(pSysLink, IID_ILinks, reinterpret_cast<IUnknown**>(&pLinks));
			return pLinks;
		};

		/// \brief <em>Creates a \c Links object</em>
		///
		/// \overload
		///
		/// Creates a \c Links object that represents a collection of hyper links and returns its implementation
		/// of a given interface as a raw pointer.
		///
		/// \param[in] pSysLink The \c SysLink object the \c Links object will work on.
		/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
		///            be returned.
		/// \param[out] ppLinks Receives the object's implementation of the interface identified by
		///             \c requiredInterface.
		static inline void InitLinks(SysLink* pSysLink, REFIID requiredInterface, IUnknown** ppLinks)
		{
			ATLASSERT_POINTER(ppLinks, IUnknown*);
			ATLASSUME(ppLinks);

			*ppLinks = NULL;
			// create a Links object and initialize it with the given parameters
			CComObject<Links>* pLinksObj = NULL;
			CComObject<Links>::CreateInstance(&pLinksObj);
			pLinksObj->AddRef();

			pLinksObj->SetOwner(pSysLink);

			pLinksObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppLinks));
			pLinksObj->Release();
		};
	#endif

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};
};     // ClassFactory