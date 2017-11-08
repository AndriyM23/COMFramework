#include "Utils.h"
#include "COMComponentKit.h"

//The header for _ASSERT() and _ASSERTE() macros of Run-Time Library.
#include  <CRTDBG.H>
#include  <CASSERT>

#include  <ALGORITHM>   // For std::for_each().

using namespace std;
using namespace Utils;

namespace COMLibrary
{

static const ULONG INVALID_REFCOUNT_VALUE{ 0xFFFFFFFFUL };

////////////////////////////////////////////////////////////////////////////////////////////////////////////////

SubComponentBase::SubComponentBase(REFIID riid)
    : iid(riid)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    InitializeCriticalSectionAndSpinCount(&csInitialization, 4000);
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
}

SubComponentBase::~SubComponentBase()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    DeleteCriticalSection(&csInitialization);
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
}

HRESULT SubComponentBase::QueryInterface(REFIID riid, void** ppvObject)
{
    trace(_TEXT(__FUNCTION__));
    assert(delegateOfContainingCOMComponent != nullptr);
    return delegateOfContainingCOMComponent->QueryInterface(riid, ppvObject);
}

ULONG SubComponentBase::AddRef(void)
{
    if (bIgnore_AddRef_Release_Calls)
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|Call ignored"));
        return 0;
    }

    assert(delegateOfContainingCOMComponent != nullptr);

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    LONG refCount_new = InterlockedIncrement(&refCount);

    InitializeSubComponent();

    ULONG refCount_fromDelegate = delegateOfContainingCOMComponent->AddRef();
    refCount_fromDelegate = refCount_fromDelegate == INVALID_REFCOUNT_VALUE ? refCount_new : refCount_fromDelegate;
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d, refCount_fromDelegate=%u"), refCount_new, bInitialized, refCount_fromDelegate);

    return refCount_fromDelegate;
}

ULONG SubComponentBase::Release(void)
{
    if (bIgnore_AddRef_Release_Calls)
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|Call ignored"));
        return 0;
    }

    assert(delegateOfContainingCOMComponent != nullptr);

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    long refCount_new = InterlockedDecrement(&refCount);
    bool bInitialized_local = bInitialized;
    ULONG refCount_fromDelegate;
    if (refCount_new == 0)
    {
        refCount_fromDelegate = delegateOfContainingCOMComponent->Release();
        ReleaseSubComponent();
        bInitialized_local = false;
    }
    else
        refCount_fromDelegate = delegateOfContainingCOMComponent->Release();

    refCount_fromDelegate = refCount_fromDelegate == INVALID_REFCOUNT_VALUE ? refCount_new  : refCount_fromDelegate;
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d, refCount_fromDelegate=%u"), refCount_new, bInitialized_local, refCount_fromDelegate);
    return refCount_fromDelegate;
}

HRESULT SubComponentBase::QueryInterface_nonDelegating(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|Calling delegateOfContainingCOMComponent->QueryInterface_nonDelegating()"));
    assert(delegateOfContainingCOMComponent != nullptr);
    return delegateOfContainingCOMComponent->QueryInterface_nonDelegating(riid, ppvObject);
}

ULONG SubComponentBase::AddRef_nonDelegating(void)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    LONG refCount_new = InterlockedIncrement(&refCount);
    InitializeSubComponent();
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d"), refCount_new, bInitialized);
    return refCount_new;
}

ULONG SubComponentBase::Release_nonDelegating(void)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    long refCount_new = InterlockedDecrement(&refCount);
    bool bInitialized_local = bInitialized;
    if (refCount_new == 0)
    {
        ReleaseSubComponent();
        bInitialized_local = false;
    }
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d"), refCount_new, bInitialized_local);
    return refCount_new;
}

void SubComponentBase::OnFirstAddRefInitialized()
{
    trace(_TEXT(__FUNCTION__));
}

void SubComponentBase::OnLastReleaseUninitialized()
{
    trace(_TEXT(__FUNCTION__));
}

HRESULT SubComponentBase::PostConstructorInitialize()
{
    trace(_TEXT(__FUNCTION__));
    return S_OK;
}

HRESULT SubComponentBase::PreDestructorRelease()
{
    trace(_TEXT(__FUNCTION__));
    return S_OK;
}

REFIID SubComponentBase::GetIID()
{
    return iid;
}

void SubComponentBase::InitializeSubComponent()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);

    EnterCriticalSection(&csInitialization);
    if (bInitialized)
    {
        LeaveCriticalSection(&csInitialization);
        trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d|Subcomponent is already initialized"), refCount, bInitialized);
        return;
    }

    //Notify superclass of enclosing component, that one of it's subcomponents has initialized itself.
    delegateOfContainingCOMComponent->OnSubComponentInitialized(this);

    // Give our subclass a chance to initialize itself.
    OnFirstAddRefInitialized();

    bInitialized = true;
    LeaveCriticalSection(&csInitialization);

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d"), refCount, bInitialized);
}

void SubComponentBase::ReleaseSubComponent()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|refCount=%d, bInitialized=%d"), refCount, bInitialized);
    assert(refCount == 0);
    EnterCriticalSection(&csInitialization);
    if (!bInitialized)
    {
        LeaveCriticalSection(&csInitialization);
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d|Subcomponent is already released."), refCount, bInitialized);
        return;
    }

    // Give our subclass a chance to free acquired resources.
    OnLastReleaseUninitialized();

    bInitialized = false;
    LeaveCriticalSection(&csInitialization);

    //Notify superclass of enclosing component, that one of it's subcomponents has uninitialized itself.
    delegateOfContainingCOMComponent->OnSubComponentReleased(this);

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|refCount=%d, bInitialized=%d"), refCount, bInitialized);
}

bool SubComponentBase::DecreaseRefCounterTo1IfNot0()
{
    if (refCount)
    {
        refCount = 1;
        return true;
    }
    return false;
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// {CE2DBFC8-B1AF-4E99-9934-7D143863DD72}
const IID COMComponentBase::UNINITIALIZED_PRIMARY_IID { 0xce2dbfc8, 0xb1af, 0x4e99,{ 0x99, 0x34, 0x7d, 0x14, 0x38, 0x63, 0xdd, 0x72 } };

COMComponentBase::COMComponentBase(REFCLSID clsid, PLONG pcntrComponentsInServer, IUnknown* iUnknownOuter)
    :   clsid(clsid),
        primaryIID(UNINITIALIZED_PRIMARY_IID),
        iUnknownOuter(iUnknownOuter),
        pcntrComponentsInServer(pcntrComponentsInServer)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));

    InitializeCriticalSectionAndSpinCount(&csInitialization, 4000);

    if (iUnknownOuter == nullptr)
        // If we are not aggregated - outer IUnknown pointer will show onto us.
        this->iUnknownOuter = reinterpret_cast<IUnknown*>(dynamic_cast<IUnknown_nonDelegating*>(this));

    if (this->pcntrComponentsInServer)
    {
        LONG new_g_cComponents = InterlockedIncrement(this->pcntrComponentsInServer);
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|g_cComponents=%d"), new_g_cComponents);
    }
    else
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
}

COMComponentBase::~COMComponentBase()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    assert(cSubComponents == 0);
    PreDestructorRelease();
    if (pcntrComponentsInServer)
    {
        LONG new_g_cComponents = InterlockedDecrement(pcntrComponentsInServer);
        trace(_TEXT(__FUNCTION__) _TEXT("|g_cComponents=%d"), new_g_cComponents);
    }

    DeleteCriticalSection(&csInitialization);
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
}

HRESULT COMComponentBase::QueryInterface_nonDelegating(REFIID riid, void **ppvObject)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    REFIID riid_to_find = riid == IID_IUnknown ? primaryIID : riid;
    DICT_SUBCOMPONENTS::const_iterator posSubComponent = subcomponents.find(riid_to_find);
    assert( !((riid == IID_IUnknown) && (posSubComponent == subcomponents.end())) );
    if (posSubComponent == subcomponents.end())
    {
        if(iUnknownInner)
        {
            // We aggregate a COM object.
            DICT_AGGREGATED_INTERFACES::const_iterator posUnknAggrInnerInterf = iUnknownInnerAggregatedInterfaces.find(riid_to_find);
            if (posUnknAggrInnerInterf != iUnknownInnerAggregatedInterfaces.end())
            {
                // This is the call to QueryInterface_nonDelegating() of inner aggregated component, actually.
                HRESULT result = iUnknownInner->QueryInterface(riid_to_find, ppvObject);
                trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Requested aggregated inner interface %s = %p."), GuidToString(riid_to_find).c_str(), *ppvObject);
                //We return here at once, because reference via inner pointer has already increased outer reference counter.
                return result;
            }
        }

        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Interface %s not found."), GuidToString(riid_to_find).c_str());
        // We don't aggregate any COM object. So, our COM component doesn't implement the requested interface.
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    if (riid == IID_IUnknown)
        *ppvObject = reinterpret_cast<IUnknown*>(dynamic_cast<IUnknown_nonDelegating*>(posSubComponent->second.get()));
    else
        *ppvObject = posSubComponent->second.get()->TypecastSubComponentToCOMInterface();

    trace(_TEXT(__FUNCTION__) _TEXT("|Requested IID=%s = %p"), GuidToString(riid_to_find).c_str(), *ppvObject);
    static_cast<IUnknown*>(*ppvObject)->AddRef(); //If IID_IUnknown was requested - we call here actually AddRef_nonDelegating() of the SubComponent.
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
    return S_OK;
}

ULONG COMComponentBase::AddRef_nonDelegating(void)
{
    trace(_TEXT(__FUNCTION__)  _TEXT("|>>>|Ret:INVALID_REFCOUNT_VALUE=0xFFFFFFFF"));
    return INVALID_REFCOUNT_VALUE;
}

ULONG COMComponentBase::Release_nonDelegating(void)
{
    trace(_TEXT(__FUNCTION__)  _TEXT("|>>>|Ret:INVALID_REFCOUNT_VALUE=0xFFFFFFFF"));
    return INVALID_REFCOUNT_VALUE;
}

HRESULT COMComponentBase::QueryInterface(REFIID riid, void **ppvObject)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    HRESULT result = iUnknownOuter->QueryInterface(riid, ppvObject);
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh"), result);
    return result;
}

ULONG COMComponentBase::AddRef(void)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    ULONG newRefCtr = iUnknownOuter->AddRef();
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|newRefCtr=%u"), newRefCtr);
    return newRefCtr;
}

ULONG COMComponentBase::Release(void)
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    ULONG newRefCtr = iUnknownOuter->Release();
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|newRefCtr=%u"), newRefCtr);
    return newRefCtr;
}

void COMComponentBase::OnSubComponentInitialized(PSubComponentBase subComponent)
{
    UNREFERENCED_PARAMETER(subComponent);

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    LONG cSubComponents_new = InterlockedIncrement(&cSubComponents);

    EnterCriticalSection(&csInitialization);
    if (!bInitialized)
    {
        // Give our subclass a chance to initialize itself.
        OnComponentInitialized(subComponent);
        bInitialized = true;
    }
    LeaveCriticalSection(&csInitialization);

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|cSubComponents_new=%d, bInitialized=%d"), cSubComponents_new, bInitialized);
}

void COMComponentBase::OnSubComponentReleased(PSubComponentBase subComponent)
{
    UNREFERENCED_PARAMETER(subComponent);

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    LONG cSubComponents_new = InterlockedDecrement(&cSubComponents);
    trace(_TEXT(__FUNCTION__) _TEXT("|cSubComponents=%d"), cSubComponents_new);
    if (cSubComponents_new == 0)
    {
        EnterCriticalSection(&csInitialization);
        if (bInitialized)
        {
            // Give our subclass a chance to free acquired resources.
            OnComponentUninitialized(subComponent);
            bInitialized = false;
        }
        LeaveCriticalSection(&csInitialization);

        delete this;
    }
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|cSubComponents=%d"), cSubComponents_new);
}

void COMComponentBase::OnComponentInitialized(PSubComponentBase subComponent)
{
    trace(_TEXT(__FUNCTION__));
    UNREFERENCED_PARAMETER(subComponent);
}

void COMComponentBase::OnComponentUninitialized(PSubComponentBase subComponent)
{
    trace(_TEXT(__FUNCTION__));
    UNREFERENCED_PARAMETER(subComponent);
}

void COMComponentBase::AddSubComponent(SubComponentBasePtr subComponent, bool bPrimarySubcomponent)
{
    if (bPrimarySubcomponent)
        primaryIID = subComponent->GetIID();
    else if (primaryIID == UNINITIALIZED_PRIMARY_IID)
        primaryIID = subComponent->GetIID();
    subComponent->delegateOfContainingCOMComponent = this;
    subcomponents.insert(make_pair(subComponent->GetIID(), subComponent));
}

HRESULT COMComponentBase::AddAggregatedComponent(REFCLSID aggregatedCLSID, const std::vector<IID>& aggregatedIIDs, DWORD dwClsContext)
{
    trace(_TEXT(__FUNCTION__));

    PSubComponentBase pPrimarySubComponent= GetPrimarySubcomponent();
    IUnknown* pUnknownOuter = static_cast<IUnknown*>(pPrimarySubComponent);
    HRESULT result = ::CoCreateInstance(aggregatedCLSID,
                                        pUnknownOuter,
                                        dwClsContext,
                                        IID_IUnknown,
                                        reinterpret_cast<void**>(&iUnknownInner));
    if (FAILED(result))
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|CoCreateInstance(clsid=%s,,,IID_IUnknown,) failed. Result=%08.8Xh."), GuidToString(aggregatedCLSID).c_str(), result);
        iUnknownInner = nullptr;
        return result;
    }

    // If after retrieving non-delegating IUnknown from aggreagted component, we have our reference counter >= 1, then we have aggregated a COM component,
    // which in turn, is also aggregating another COM component. We decrease this reference counter to 1, and remember, that we have aggregated a complex COM component.
    // If we have aggregated a simple COM object, our reference counter will be 0, and we don't change it.
    // See function COMComponentBase::PostConstructorInitialize() for further usage of bAggregatedComponentIsAlsoAggregating flag.
    bAggregatedComponentIsAlsoAggregating = pPrimarySubComponent->DecreaseRefCounterTo1IfNot0();
    if (bAggregatedComponentIsAlsoAggregating)
        trace(_TEXT(__FUNCTION__) _TEXT("|We aggregate a COM component, which is also aggregating."));
    
    trace(_TEXT(__FUNCTION__) _TEXT("|Get reference to internal interfaces of aggregated inner component."));

    for (std::vector<IID>::const_iterator citerAggriids = aggregatedIIDs.cbegin(); citerAggriids != aggregatedIIDs.cend(); ++citerAggriids)
    {
        IUnknown* iUnknownAggregated = nullptr;
        result = iUnknownInner->QueryInterface(*citerAggriids, reinterpret_cast<void**>(&iUnknownAggregated));
        if(FAILED(result))
        {
            trace(_TEXT(__FUNCTION__) _TEXT("|Inner component does not support interface %s. Result=%08.8Xh."), GuidToString(*citerAggriids).c_str(), result);

            // Perform cleanup of previously obtained inner interfaces and iUnknown inner pointer.
            PSubComponentBase primarySubComponent = GetPrimarySubcomponent();
            primarySubComponent->bIgnore_AddRef_Release_Calls = true;
            for_each(iUnknownInnerAggregatedInterfaces.begin(), iUnknownInnerAggregatedInterfaces.end(), [](const PAIR_GUID_IUNKNOWN& iUnknownInnerPair)
            {
                iUnknownInnerPair.second->Release();
            });
            primarySubComponent->bIgnore_AddRef_Release_Calls = false;

            iUnknownInnerAggregatedInterfaces.clear();
            iUnknownInner->Release();   //calling actually m_pUnknownInner->Release_nonDelegating();
            iUnknownInner = nullptr;

            if (bAggregatedComponentIsAlsoAggregating)
                pPrimarySubComponent->refCount = 0;

            break;
        }
        else
        {
            iUnknownInnerAggregatedInterfaces.insert(make_pair(*citerAggriids, iUnknownAggregated));
            trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Got reference (%p) to interface %s of inner component"), iUnknownAggregated, GuidToString(*citerAggriids).c_str());
        }
    }
    return result;
}

PSubComponentBase COMComponentBase::GetPrimarySubcomponent()
{
    DICT_SUBCOMPONENTS::const_iterator posPrimarySubComponent = subcomponents.find(primaryIID);
    return static_cast<PSubComponentBase>(posPrimarySubComponent->second.get());
}

REFCLSID COMComponentBase::GetCLSID()
{
    return clsid;
}

HRESULT COMComponentBase::PostConstructorInitialize()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    HRESULT result = S_OK;
    if (!bPostConstructorInitDone)
    {
        std::for_each(subcomponents.begin(), subcomponents.end(), [&result](const PAIR_GUID_SUBCOMPONENT& subcomponentPair)
        {
            HRESULT result_local = subcomponentPair.second->PostConstructorInitialize();
            if (FAILED(result_local))
                result = result_local;
        });

        if (iUnknownInner)
        {
            // We are aggregating. All aggregated interfaces now have to have taken reference on us. 
            // We are going to dispose of all those references, but without invoking Release() of aggregared interfaces.
            IUnknown* iUnknownPrimary = static_cast<IUnknown*>(GetPrimarySubcomponent());
            // If we are aggregating a COM component, which in turn is also aggregating another COM component,
            // then our reference counter is already 1, which was incremented during obtaining non-delegating IUnknown of the inner(aggregated) component.
            // So, we have to release all of the references, taken by aggregated interfaces.
            // Otherwise, if inner component is not aggregating, we have to release on one reference fewer, 
            // than the amount of aggregated interfaces - so we will not delete ourself.
            for (size_t i = bAggregatedComponentIsAlsoAggregating ? 0 : 1; i<iUnknownInnerAggregatedInterfaces.size(); ++i)
                iUnknownPrimary->Release();
        }

        bPostConstructorInitDone = true;
    }
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
    return result;
}

HRESULT COMComponentBase::PreDestructorRelease()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    HRESULT result = S_OK;
    if (bPostConstructorInitDone)
    {
        if (iUnknownInner)
        {
            // We must release all aggregated interfaces. 
            PSubComponentBase primarySubComponent = GetPrimarySubcomponent();
            primarySubComponent->bIgnore_AddRef_Release_Calls = true;
            for_each(iUnknownInnerAggregatedInterfaces.begin(), iUnknownInnerAggregatedInterfaces.end(), [](const PAIR_GUID_IUNKNOWN& iUnknownInnerPair)
            {
                iUnknownInnerPair.second->Release();
            });
            primarySubComponent->bIgnore_AddRef_Release_Calls = false;

            iUnknownInnerAggregatedInterfaces.clear();
            iUnknownInner->Release();
            iUnknownInner = nullptr;
        }

        for_each(subcomponents.rbegin(), subcomponents.rend(), [&result](const PAIR_GUID_SUBCOMPONENT& subcomponentPair)
        {
            HRESULT result_local = subcomponentPair.second->PreDestructorRelease();
            if (FAILED(result_local))
                result = result_local;
        });

        bPostConstructorInitDone = false;
    }
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
    return result;
}

}