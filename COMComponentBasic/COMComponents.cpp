#pragma once

#include "Utils.h"
#define DECLARE_AND_DEFINE_GUIDS
#include "COMComponents.h"

using namespace std;
using namespace Utils;
using namespace COMLibrary;

SubComponentInterf01::SubComponentInterf01()
    : COMLibrary::SubComponentBase(IID_IInterfBasic01)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf01::F01_IInterfBasic01()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_S_SCCESS1;
}

HRESULT SubComponentInterf01::F02_IInterfBasic01()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_E_ERROR1;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf01, IInterfBasic01)

//===================================================================================================

SubComponentInterf01Proxy::SubComponentInterf01Proxy()
    : COMLibrary::SubComponentBase(IID_IInterfBasic01)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf01Proxy::F01_IInterfBasic01()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return includedInterf01->F01_IInterfBasic01();
}

HRESULT SubComponentInterf01Proxy::F02_IInterfBasic01()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return includedInterf01->F02_IInterfBasic01();
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf01Proxy, IInterfBasic01)

HRESULT SubComponentInterf01Proxy::PostConstructorInitialize()
{
    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));

    HRESULT result = ::CoCreateInstance(CLSID_COMComponent01,
                                        nullptr,
                                        CLSCTX_INPROC_SERVER,
                                        includedInterf01.GetIID(),
                                        reinterpret_cast<LPVOID*>(&includedInterf01));

    if (FAILED(result))
    {
        Utils::trace(_TEXT(__FUNCTION__) _TEXT("|>>>|CoCreateInstance(CLSID_COMComponent01,,,IID_IInterfBasic01,) failed. Result=%08.8Xh."), result);
        return result;
    }

    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    return result;
}

#if 0
// The smart pointer will be resetted in any case during deletion. So, there is no sense here to overload this function.
HRESULT SubComponentInterf01Proxy::PreDestructorRelease()
{
    Utils::trace(_TEXT(__FUNCTION__));
    includedInterf01.Reset();
    return S_OK;
}
#endif

//===================================================================================================

SubComponentInterf02::SubComponentInterf02()
    : COMLibrary::SubComponentBase(IID_IInterfBasic02)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf02::F01_IInterfBasic02()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_S_SCCESS2;
}

HRESULT SubComponentInterf02::F02_IInterfBasic02()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_E_ERROR2;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf02, IInterfBasic02)

//===================================================================================================

SubComponentInterf03::SubComponentInterf03()
    : COMLibrary::SubComponentBase(IID_IInterfBasic03)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf03::F01_IInterfBasic03()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_S_SCCESS3;
}

HRESULT SubComponentInterf03::F02_IInterfBasic03()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_E_ERROR3;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf03, IInterfBasic03)

//===================================================================================================

SubComponentInterf04::SubComponentInterf04()
    : COMLibrary::SubComponentBase(IID_IInterfBasic04)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf04::F01_IInterfBasic04()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_S_SCCESS4;
}

HRESULT SubComponentInterf04::F02_IInterfBasic04()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_E_ERROR4;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf04, IInterfBasic04)

//===================================================================================================

SubComponentInterf04Proxy::SubComponentInterf04Proxy()
    : COMLibrary::SubComponentBase(IID_IInterfBasic04)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf04Proxy::F01_IInterfBasic04()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return includedInterf04->F01_IInterfBasic04();
}

HRESULT SubComponentInterf04Proxy::F02_IInterfBasic04()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return includedInterf04->F02_IInterfBasic04();
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf04Proxy, IInterfBasic04)

HRESULT SubComponentInterf04Proxy::PostConstructorInitialize()
{
    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));

    HRESULT result = ::CoCreateInstance(CLSID_COMComponent04,
                                        nullptr,
                                        CLSCTX_INPROC_SERVER,
                                        includedInterf04.GetIID(),
                                        reinterpret_cast<LPVOID*>(&includedInterf04));

    if (FAILED(result))
    {
        Utils::trace(_TEXT(__FUNCTION__) _TEXT("|>>>|CoCreateInstance(CLSID_COMComponent04,,,IID_IInterfBasic04,) failed. Result=%08.8Xh."), result);
        return result;
    }

    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    return result;
}

#if 0
// The smart pointer will be resetted in any case during deletion. So, there is no sense here to overload this function.
HRESULT SubComponentInterf04Proxy::PreDestructorRelease()
{
    Utils::trace(_TEXT(__FUNCTION__));
    includedInterf04.Reset();
    return S_OK;
}
#endif

//===================================================================================================

SubComponentInterf05::SubComponentInterf05()
    : COMLibrary::SubComponentBase(IID_IInterfBasic05)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf05::F01_IInterfBasic05()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_S_SCCESS5;
}

HRESULT SubComponentInterf05::F02_IInterfBasic05()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_E_ERROR5;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf05, IInterfBasic05)

//===================================================================================================

SubComponentInterf06::SubComponentInterf06()
    : COMLibrary::SubComponentBase(IID_IInterfBasic06)
{
    Utils::trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentInterf06::F01_IInterfBasic06()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_S_SCCESS6;
}

HRESULT SubComponentInterf06::F02_IInterfBasic06()
{
    Utils::trace(_TEXT(__FUNCTION__));
    return COMPONENTBASIC_E_ERROR6;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentInterf06, IInterfBasic06)

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

COMComponent01::COMComponent01(PLONG        pcntrComponentsInServer,   ///< Address of active components counter within the COM server (DLL or EXE).
                               IUnknown*    iUnknownOuter)
    : COMLibrary::COMComponentBase::COMComponentBase(CLSID_COMComponent01, pcntrComponentsInServer, iUnknownOuter)
{
    Utils::trace(_TEXT(__FUNCTION__));

    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentInterf01);
    AddSubComponent(subComponentPtr);
}

//===================================================================================================

COMComponent04::COMComponent04(PLONG        pcntrComponentsInServer,   ///< Address of active components counter within the COM server (DLL or EXE).
                               IUnknown*    iUnknownOuter)
    : COMLibrary::COMComponentBase::COMComponentBase(CLSID_COMComponent04, pcntrComponentsInServer, iUnknownOuter)
{
    Utils::trace(_TEXT(__FUNCTION__));

    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentInterf04);
    AddSubComponent(subComponentPtr);
}

//===================================================================================================

COMComponentI040506::COMComponentI040506(PLONG        pcntrComponentsInServer,   ///< Address of active components counter within the COM server (DLL or EXE).
                                             IUnknown*    iUnknownOuter)
    : COMLibrary::COMComponentBase::COMComponentBase(CLSID_COMComponentI040506, pcntrComponentsInServer, iUnknownOuter)
{
    Utils::trace(_TEXT(__FUNCTION__));

    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentInterf04Proxy);
    AddSubComponent(subComponentPtr, true);
    subComponentPtr.reset(new SubComponentInterf05);
    AddSubComponent(subComponentPtr);
    subComponentPtr.reset(new SubComponentInterf06);
    AddSubComponent(subComponentPtr);
}

//===================================================================================================

COMComponentI010203A04A05A06::COMComponentI010203A04A05A06(PLONG  pcntrComponentsInServer)   ///< Address of active components counter within the COM server (DLL or EXE).
    : COMLibrary::COMComponentBase::COMComponentBase(CLSID_COMComponentI010203A04A05A06, pcntrComponentsInServer)
{
    Utils::trace(_TEXT(__FUNCTION__));

    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentInterf01Proxy);
    AddSubComponent(subComponentPtr);
    subComponentPtr.reset(new SubComponentInterf02);
    AddSubComponent(subComponentPtr, true);
    subComponentPtr.reset(new SubComponentInterf03);
    AddSubComponent(subComponentPtr);
}

HRESULT COMComponentI010203A04A05A06::PostConstructorInitialize()
{
    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));

    vector<IID> aggriids { IID_IInterfBasic04, IID_IInterfBasic05, IID_IInterfBasic06 };
    HRESULT result = AddAggregatedComponent(CLSID_COMComponentI040506, aggriids);
    if (FAILED(result))
    {
        Utils::trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Failed to aggregate component CLSID_COMComponentI040506=%s. Result=%08.8Xh."), GuidToString(CLSID_COMComponentI040506).c_str(), result);
        return result;
    }

    return COMLibrary::COMComponentBase::PostConstructorInitialize();
}

//===================================================================================================

COMComponent02A01::COMComponent02A01(PLONG          pcntrComponentsInServer,    ///< Address of active components counter within the COM server (DLL or EXE).
                                     IUnknown*      iUnknownOuter )
    : COMLibrary::COMComponentBase::COMComponentBase(CLSID_COMComponent02A01, pcntrComponentsInServer, iUnknownOuter)
{
    Utils::trace(_TEXT(__FUNCTION__));

    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentInterf02);
    AddSubComponent(subComponentPtr);
}

HRESULT COMComponent02A01::PostConstructorInitialize()
{
    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));

    vector<IID> aggriids{ IID_IInterfBasic01 };
    HRESULT result = AddAggregatedComponent(CLSID_COMComponent01, aggriids);
    if (FAILED(result))
    {
        Utils::trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Failed to aggregate component CLSID_COMComponent01=%s. Result=%08.8Xh."), GuidToString(CLSID_COMComponent01).c_str(), result);
        return result;
    }

    return COMLibrary::COMComponentBase::PostConstructorInitialize();
}

//===================================================================================================

COMComponentI04A02AA01::COMComponentI04A02AA01(PLONG  pcntrComponentsInServer)   ///< Address of active components counter within the COM server (DLL or EXE).
    : COMLibrary::COMComponentBase::COMComponentBase(CLSID_COMComponentI04A02AA01, pcntrComponentsInServer)
{
    Utils::trace(_TEXT(__FUNCTION__));

    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentInterf04Proxy);
    AddSubComponent(subComponentPtr);
}

HRESULT COMComponentI04A02AA01::PostConstructorInitialize()
{
    Utils::trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));

    vector<IID> aggriids{ IID_IInterfBasic01, IID_IInterfBasic02 };
    HRESULT result = AddAggregatedComponent(CLSID_COMComponent02A01, aggriids);
    if (FAILED(result))
    {
        Utils::trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Failed to aggregate component CLSID_COMComponent02A01=%s. Result=%08.8Xh."), GuidToString(CLSID_COMComponent02A01).c_str(), result);
        return result;
    }

    return COMLibrary::COMComponentBase::PostConstructorInitialize();
}
