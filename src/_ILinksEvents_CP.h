//////////////////////////////////////////////////////////////////////
/// \class Proxy_ILinksEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ILinksEvents interface</em>
///
/// \sa Links, LblCtlsLibU::_ILinksEvents
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ILinksEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ILinksEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ILinksEvents