#include "COMComponentFactoryKit.h"
#include "Utils.h"

using namespace Utils;
using namespace std;

namespace COMLibrary
{

SubComponentFactoryBase::SubComponentFactoryBase(REFCLSID                       clsid,
                                                 PFUNCTION_CREATE_COMPONENT     fnCreateComponent,
                                                 PLONG                          pcntrServerLocks )
    : SubComponentBase(IID_IClassFactory),
      clsid(clsid),
      fnCreateComponent(fnCreateComponent),
      pcntrServerLocks(pcntrServerLocks)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|CLSID=%s"), GuidToString(clsid).c_str());
}

HRESULT SubComponentFactoryBase::CreateInstance(IUnknown*   pIUnknownOuter,
                                                REFIID      riid,
                                                void**      ppvObject)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|CLSID=%s"), GuidToString(clsid).c_str());
    HRESULT result = fnCreateComponent (pIUnknownOuter, riid, ppvObject, reinterpret_cast<PCOMComponentFactoryBase>(delegateOfContainingCOMComponent));
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|CLSID=%s|RET:%08.8Xh|*ppvObject=%p"), GuidToString(clsid).c_str(), result, *ppvObject);
    return result;
}

HRESULT SubComponentFactoryBase::LockServer(BOOL fLock)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|CLSID=%s|fLock=%d"), GuidToString(clsid).c_str(), fLock);
    LONG newCntrServerLocksVal;
    if (fLock)
        newCntrServerLocksVal = InterlockedIncrement(pcntrServerLocks);
    else
        newCntrServerLocksVal = InterlockedDecrement(pcntrServerLocks);
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|CLSID=%s|RET:S_OK|g_cServerLocks=%d"),  GuidToString(clsid).c_str(), newCntrServerLocksVal);
    return S_OK;
}

SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(SubComponentFactoryBase, IClassFactory)

//-----------------------------------------------------------------------------------------------------

COMComponentFactoryBase::COMComponentFactoryBase(REFCLSID                    clsid,
                                                 PFUNCTION_CREATE_COMPONENT  fnCreateComponent,
                                                 PLONG                       pcntrServerLocks )
    : COMLibrary::COMComponentBase::COMComponentBase(clsid)
{
    Utils::trace(_TEXT(__FUNCTION__));
    COMLibrary::SubComponentBasePtr subComponentPtr;

    subComponentPtr.reset(new SubComponentFactoryBase(clsid, fnCreateComponent, pcntrServerLocks));
    AddSubComponent(subComponentPtr);
}
}
