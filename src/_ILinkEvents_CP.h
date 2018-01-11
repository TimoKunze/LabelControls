//////////////////////////////////////////////////////////////////////
/// \class Proxy_ILinkEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _ILinkEvents interface</em>
///
/// \sa Link, LblCtlsLibU::_ILinkEvents
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_ILinkEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_ILinkEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_ILinkEvents