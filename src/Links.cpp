// Links.cpp: Manages a collection of Link objects

#include "stdafx.h"
#include "Links.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP Links::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_ILinks, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP Links::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<Links>* pLinksObj = NULL;
	CComObject<Links>::CreateInstance(&pLinksObj);
	pLinksObj->AddRef();

	// clone all settings
	properties.CopyTo(&pLinksObj->properties);

	pLinksObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pLinksObj->Release();
	return S_OK;
}

STDMETHODIMP Links::Next(ULONG numberOfMaxLinks, VARIANT* pLinks, ULONG* pNumberOfLinksReturned)
{
	ATLASSERT_POINTER(pLinks, VARIANT);
	if(!pLinks) {
		return E_POINTER;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	// check each link in the control
	ULONG i = 0;
	for(i = 0; i < numberOfMaxLinks; ++i) {
		VariantInit(&pLinks[i]);

		do {
			if(properties.lastEnumeratedLink >= 0) {
				properties.lastEnumeratedLink = GetNextLinkToProcess(properties.lastEnumeratedLink, hWndSysLink);
			} else {
				properties.lastEnumeratedLink = GetFirstLinkToProcess(hWndSysLink);
			}
		} while((properties.lastEnumeratedLink != -1) && (!IsPartOfCollection(properties.lastEnumeratedLink, hWndSysLink)));

		if(properties.lastEnumeratedLink != -1) {
			ClassFactory::InitLink(properties.lastEnumeratedLink, properties.pOwnerSysLink, IID_IDispatch, reinterpret_cast<LPUNKNOWN*>(&pLinks[i].pdispVal));
			pLinks[i].vt = VT_DISPATCH;
		} else {
			// there's nothing more to iterate
			break;
		}
	}
	if(pNumberOfLinksReturned) {
		*pNumberOfLinksReturned = i;
	}

	return (i == numberOfMaxLinks ? S_OK : S_FALSE);
}

STDMETHODIMP Links::Reset(void)
{
	properties.lastEnumeratedLink = -1;
	return S_OK;
}

STDMETHODIMP Links::Skip(ULONG numberOfLinksToSkip)
{
	VARIANT dummy;
	ULONG numLinksReturned = 0;
	// we could skip all links at once, but it's easier to skip them one after the other
	for(ULONG i = 1; i <= numberOfLinksToSkip; ++i) {
		HRESULT hr = Next(1, &dummy, &numLinksReturned);
		VariantClear(&dummy);
		if(hr != S_OK || numLinksReturned == 0) {
			// there're no more links to skip, so don't try it anymore
			break;
		}
	}
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


BOOL Links::IsValidBooleanFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BOOL);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL Links::IsValidIntegerFilter(VARIANT& filter, int minValue)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT)) && element.intVal >= minValue;
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

BOOL Links::IsValidStringFilter(VARIANT& filter)
{
	BOOL isValid = TRUE;

	if((filter.vt == (VT_ARRAY | VT_VARIANT)) && filter.parray) {
		LONG lBound = 0;
		LONG uBound = 0;

		if((SafeArrayGetLBound(filter.parray, 1, &lBound) == S_OK) && (SafeArrayGetUBound(filter.parray, 1, &uBound) == S_OK)) {
			// now that we have the bounds, iterate the array
			VARIANT element;
			if(lBound > uBound) {
				isValid = FALSE;
			}
			for(LONG i = lBound; i <= uBound && isValid; ++i) {
				if(SafeArrayGetElement(filter.parray, &i, &element) == S_OK) {
					isValid = (element.vt == VT_BSTR);
					VariantClear(&element);
				} else {
					isValid = FALSE;
				}
			}
		} else {
			isValid = FALSE;
		}
	} else {
		isValid = FALSE;
	}

	return isValid;
}

int Links::GetFirstLinkToProcess(HWND hWndSysLink)
{
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = 0;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return link.iLink;
	}
	return -1;
}

int Links::GetNextLinkToProcess(int previousLink, HWND hWndSysLink)
{
	ATLASSERT(IsWindow(hWndSysLink));

	LITEM link = {0};
	link.iLink = previousLink + 1;
	link.mask = LIF_ITEMINDEX | LIF_STATE;
	if(SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
		return link.iLink;
	}
	return -1;
}

BOOL Links::IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(VARIANT_BOOL, VARIANT_BOOL);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.boolVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(element.boolVal == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL Links::IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		int v = 0;
		if(SUCCEEDED(VariantChangeType(&element, &element, 0, VT_INT))) {
			v = element.intVal;
		}
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(LONG, LONG);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, v)) {
				foundMatch = TRUE;
			}
		} else {
			if(v == value) {
				foundMatch = TRUE;
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL Links::IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction)
{
	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(pSafeArray, 1, &lBound);
	SafeArrayGetUBound(pSafeArray, 1, &uBound);

	VARIANT element;
	BOOL foundMatch = FALSE;
	for(LONG i = lBound; i <= uBound && !foundMatch; ++i) {
		SafeArrayGetElement(pSafeArray, &i, &element);
		if(pComparisonFunction) {
			typedef BOOL WINAPI ComparisonFn(BSTR, BSTR);
			ComparisonFn* pComparisonFn = reinterpret_cast<ComparisonFn*>(pComparisonFunction);
			if(pComparisonFn(value, element.bstrVal)) {
				foundMatch = TRUE;
			}
		} else {
			if(properties.caseSensitiveFilters) {
				if(lstrcmpW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			} else {
				if(lstrcmpiW(OLE2W(element.bstrVal), OLE2W(value)) == 0) {
					foundMatch = TRUE;
				}
			}
		}
		VariantClear(&element);
	}

	return foundMatch;
}

BOOL Links::IsPartOfCollection(int linkIndex, HWND hWndSysLink/* = NULL*/)
{
	if(!hWndSysLink) {
		hWndSysLink = properties.GetSysLinkHWnd();
	}
	ATLASSERT(IsWindow(hWndSysLink));

	if(!IsValidLinkIndex(linkIndex, hWndSysLink)) {
		return FALSE;
	}

	BOOL isPart = FALSE;
	// we declare this one here to avoid compiler warnings
	LITEM link = {0};
	link.iLink = linkIndex;
	link.mask = LIF_ITEMINDEX;

	BOOL mustRetrieveLinkData = FALSE;
	CComBSTR text;
	if(properties.filterType[fpIndex] != ftDeactivated) {
		if(IsIntegerInSafeArray(properties.filter[fpIndex].parray, linkIndex, properties.comparisonFunction[fpIndex])) {
			if(properties.filterType[fpIndex] == ftExcluding) {
				goto Exit;
			}
		} else {
			if(properties.filterType[fpIndex] == ftIncluding) {
				goto Exit;
			}
		}
	}

	if(properties.filterType[fpID] != ftDeactivated) {
		link.mask |= LIF_ITEMID;
		ZeroMemory(link.szID, MAX_LINKID_TEXT * sizeof(TCHAR));
		mustRetrieveLinkData = TRUE;
	}
	if(properties.filterType[fpText] != ftDeactivated) {
		CComPtr<ILink> pLink = ClassFactory::InitLink(linkIndex, properties.pOwnerSysLink);
		if(pLink) {
			pLink->get_Text(&text);
		}
	}
	if(properties.filterType[fpDrawAsNormalText] != ftDeactivated || properties.filterType[fpVisited] != ftDeactivated || properties.filterType[fpHot] != ftDeactivated || properties.filterType[fpEnabled] != ftDeactivated) {
		link.mask |= LIF_STATE;
		link.stateMask = LIS_DEFAULTCOLORS | LIS_ENABLED | LIS_HOTTRACK | LIS_VISITED;
		mustRetrieveLinkData = TRUE;
	}
	if(properties.filterType[fpURL] != ftDeactivated) {
		link.mask |= LIF_URL;
		ZeroMemory(link.szUrl, L_MAX_URL_LENGTH * sizeof(TCHAR));
		mustRetrieveLinkData = TRUE;
	}

	if(mustRetrieveLinkData) {
		if(!SendMessage(hWndSysLink, LM_GETITEM, 0, reinterpret_cast<LPARAM>(&link))) {
			goto Exit;
		}

		// apply the filters

		if(properties.filterType[fpID] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpID].parray, CComBSTR(link.szID), properties.comparisonFunction[fpID])) {
				if(properties.filterType[fpID] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpID] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpText] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpText].parray, text, properties.comparisonFunction[fpText])) {
				if(properties.filterType[fpText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpText] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpURL] != ftDeactivated) {
			if(IsStringInSafeArray(properties.filter[fpURL].parray, CComBSTR(link.szUrl), properties.comparisonFunction[fpURL])) {
				if(properties.filterType[fpURL] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpURL] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpVisited] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpVisited].parray, BOOL2VARIANTBOOL(link.stateMask & LIS_VISITED), properties.comparisonFunction[fpVisited])) {
				if(properties.filterType[fpVisited] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpVisited] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpHot] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpHot].parray, BOOL2VARIANTBOOL(link.stateMask & LIS_HOTTRACK), properties.comparisonFunction[fpHot])) {
				if(properties.filterType[fpHot] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpHot] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpEnabled] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpEnabled].parray, BOOL2VARIANTBOOL(link.stateMask & LIS_ENABLED), properties.comparisonFunction[fpEnabled])) {
				if(properties.filterType[fpEnabled] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpEnabled] == ftIncluding) {
					goto Exit;
				}
			}
		}

		if(properties.filterType[fpDrawAsNormalText] != ftDeactivated) {
			if(IsBooleanInSafeArray(properties.filter[fpDrawAsNormalText].parray, BOOL2VARIANTBOOL(link.stateMask & LIS_DEFAULTCOLORS), properties.comparisonFunction[fpDrawAsNormalText])) {
				if(properties.filterType[fpDrawAsNormalText] == ftExcluding) {
					goto Exit;
				}
			} else {
				if(properties.filterType[fpDrawAsNormalText] == ftIncluding) {
					goto Exit;
				}
			}
		}
	}
	isPart = TRUE;

Exit:
	return isPart;
}

void Links::OptimizeFilter(FilteredPropertyConstants filteredProperty)
{
	if(filteredProperty != fpDrawAsNormalText && filteredProperty != fpVisited && filteredProperty != fpHot && filteredProperty != fpEnabled) {
		// currently we optimize boolean filters only
		return;
	}

	LONG lBound = 0;
	LONG uBound = 0;
	SafeArrayGetLBound(properties.filter[filteredProperty].parray, 1, &lBound);
	SafeArrayGetUBound(properties.filter[filteredProperty].parray, 1, &uBound);

	// now that we have the bounds, iterate the array
	VARIANT element;
	UINT numberOfTrues = 0;
	UINT numberOfFalses = 0;
	for(LONG i = lBound; i <= uBound; ++i) {
		SafeArrayGetElement(properties.filter[filteredProperty].parray, &i, &element);
		if(element.boolVal == VARIANT_FALSE) {
			++numberOfFalses;
		} else {
			++numberOfTrues;
		}

		VariantClear(&element);
	}

	if(numberOfTrues > 0 && numberOfFalses > 0) {
		// we've something like True Or False Or True - we can deactivate this filter
		properties.filterType[filteredProperty] = ftDeactivated;
		VariantClear(&properties.filter[filteredProperty]);
	} else if(numberOfTrues == 0 && numberOfFalses > 1) {
		// False Or False Or False... is still False, so we need just one False
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_FALSE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	} else if(numberOfFalses == 0 && numberOfTrues > 1) {
		// True Or True Or True... is still True, so we need just one True
		VariantClear(&properties.filter[filteredProperty]);
		properties.filter[filteredProperty].vt = VT_ARRAY | VT_VARIANT;
		properties.filter[filteredProperty].parray = SafeArrayCreateVectorEx(VT_VARIANT, 1, 1, NULL);

		VARIANT newElement;
		VariantInit(&newElement);
		newElement.vt = VT_BOOL;
		newElement.boolVal = VARIANT_TRUE;
		LONG index = 1;
		SafeArrayPutElement(properties.filter[filteredProperty].parray, &index, &newElement);
		VariantClear(&newElement);
	}
}


Links::Properties::~Properties()
{
	for(int i = 0; i < NUMBEROFFILTERS_SL; ++i) {
		VariantClear(&filter[i]);
	}
	if(pOwnerSysLink) {
		pOwnerSysLink->Release();
	}
}

void Links::Properties::CopyTo(Links::Properties* pTarget)
{
	ATLASSERT_POINTER(pTarget, Properties);
	if(pTarget) {
		pTarget->pOwnerSysLink = this->pOwnerSysLink;
		if(pTarget->pOwnerSysLink) {
			pTarget->pOwnerSysLink->AddRef();
		}
		pTarget->lastEnumeratedLink = this->lastEnumeratedLink;
		pTarget->caseSensitiveFilters = this->caseSensitiveFilters;

		for(int i = 0; i < NUMBEROFFILTERS_SL; ++i) {
			VariantCopy(&pTarget->filter[i], &this->filter[i]);
			pTarget->filterType[i] = this->filterType[i];
			pTarget->comparisonFunction[i] = this->comparisonFunction[i];
		}
		pTarget->usingFilters = this->usingFilters;
	}
}

HWND Links::Properties::GetSysLinkHWnd(void)
{
	ATLASSUME(pOwnerSysLink);

	OLE_HANDLE handle = NULL;
	pOwnerSysLink->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void Links::SetOwner(SysLink* pOwner)
{
	if(properties.pOwnerSysLink) {
		properties.pOwnerSysLink->Release();
	}
	properties.pOwnerSysLink = pOwner;
	if(properties.pOwnerSysLink) {
		properties.pOwnerSysLink->AddRef();
	}
}


STDMETHODIMP Links::get_CaseSensitiveFilters(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.caseSensitiveFilters);
	return S_OK;
}

STDMETHODIMP Links::put_CaseSensitiveFilters(VARIANT_BOOL newValue)
{
	properties.caseSensitiveFilters = VARIANTBOOL2BOOL(newValue);
	return S_OK;
}

STDMETHODIMP Links::get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpDrawAsNormalText:
		case fpIndex:
		case fpVisited:
		case fpText:
		case fpHot:
		case fpURL:
		case fpID:
		case fpEnabled:
			*pValue = static_cast<LONG>(reinterpret_cast<LONG_PTR>(properties.comparisonFunction[filteredProperty]));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP Links::put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue)
{
	switch(filteredProperty) {
		case fpDrawAsNormalText:
		case fpIndex:
		case fpVisited:
		case fpText:
		case fpHot:
		case fpURL:
		case fpID:
		case fpEnabled:
			properties.comparisonFunction[filteredProperty] = reinterpret_cast<LPVOID>(static_cast<LONG_PTR>(newValue));
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP Links::get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpDrawAsNormalText:
		case fpIndex:
		case fpVisited:
		case fpText:
		case fpHot:
		case fpURL:
		case fpID:
		case fpEnabled:
			VariantClear(pValue);
			VariantCopy(pValue, &properties.filter[filteredProperty]);
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP Links::put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue)
{
	// check 'newValue'
	switch(filteredProperty) {
		case fpIndex:
			if(!IsValidIntegerFilter(newValue, 0)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpID:
		case fpText:
		case fpURL:
			if(!IsValidStringFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		case fpDrawAsNormalText:
		case fpVisited:
		case fpHot:
		case fpEnabled:
			if(!IsValidBooleanFilter(newValue)) {
				// invalid value - raise VB runtime error 380
				return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			}
			break;
		default:
			return E_INVALIDARG;
			break;
	}

	VariantClear(&properties.filter[filteredProperty]);
	VariantCopy(&properties.filter[filteredProperty], &newValue);
	OptimizeFilter(filteredProperty);
	return S_OK;
}

STDMETHODIMP Links::get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, FilterTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(filteredProperty) {
		case fpDrawAsNormalText:
		case fpIndex:
		case fpVisited:
		case fpText:
		case fpHot:
		case fpURL:
		case fpID:
		case fpEnabled:
			*pValue = properties.filterType[filteredProperty];
			return S_OK;
			break;
	}
	return E_INVALIDARG;
}

STDMETHODIMP Links::put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue)
{
	if(newValue < 0 || newValue > 2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	BOOL isOkay = FALSE;
	switch(filteredProperty) {
		case fpDrawAsNormalText:
		case fpIndex:
		case fpVisited:
		case fpText:
		case fpHot:
		case fpURL:
		case fpID:
		case fpEnabled:
			isOkay = TRUE;
			break;
	}
	if(isOkay) {
		properties.filterType[filteredProperty] = newValue;
		if(newValue != ftDeactivated) {
			properties.usingFilters = TRUE;
		} else {
			properties.usingFilters = FALSE;
			for(int i = 0; i < NUMBEROFFILTERS_SL; ++i) {
				if(properties.filterType[i] != ftDeactivated) {
					properties.usingFilters = TRUE;
					break;
				}
			}
		}
		return S_OK;
	}
	return E_INVALIDARG;
}

STDMETHODIMP Links::get_Item(LONG linkIdentifier, LinkIdentifierTypeConstants linkIdentifierType/* = litIndex*/, ILink** ppLink/* = NULL*/)
{
	ATLASSERT_POINTER(ppLink, ILink*);
	if(!ppLink) {
		return E_POINTER;
	}

	// retrieve the link's index
	int linkIndex = -1;
	switch(linkIdentifierType) {
		case litIndex:
			linkIndex = linkIdentifier;
			break;
	}

	if(linkIndex > -1) {
		if(IsPartOfCollection(linkIndex)) {
			ClassFactory::InitLink(linkIndex, properties.pOwnerSysLink, IID_ILink, reinterpret_cast<LPUNKNOWN*>(ppLink));
		}
		return S_OK;
	} else {
		// link not found
		if(linkIdentifierType == litIndex) {
			return DISP_E_BADINDEX;
		} else {
			return E_INVALIDARG;
		}
	}
}

STDMETHODIMP Links::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP Links::Contains(LONG linkIdentifier, LinkIdentifierTypeConstants linkIdentifierType/* = litIndex*/, VARIANT_BOOL* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	// retrieve the link's index
	int linkIndex = -1;
	switch(linkIdentifierType) {
		case litIndex:
			linkIndex = linkIdentifier;
			break;
	}

	*pValue = BOOL2VARIANTBOOL(IsPartOfCollection(linkIndex));
	return S_OK;
}

STDMETHODIMP Links::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndSysLink = properties.GetSysLinkHWnd();
	ATLASSERT(IsWindow(hWndSysLink));

	// count the links "by hand"
	*pValue = 0;
	int linkIndex = GetFirstLinkToProcess(hWndSysLink);
	while(linkIndex != -1) {
		if(IsPartOfCollection(linkIndex, hWndSysLink)) {
			++(*pValue);
		}
		linkIndex = GetNextLinkToProcess(linkIndex, hWndSysLink);
	}
	return S_OK;
}