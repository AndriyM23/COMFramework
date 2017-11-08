#pragma once

#include "IComPtr.h"

namespace COMLibrary
{

///////////////////////////////////////////////////////////////////////////////////////////////////////
interface IUnknown_nonDelegating
{
    virtual HRESULT STDMETHODCALLTYPE QueryInterface_nonDelegating(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject) = 0;
    virtual ULONG STDMETHODCALLTYPE AddRef_nonDelegating(void) = 0;
    virtual ULONG STDMETHODCALLTYPE Release_nonDelegating(void) = 0;
};
typedef IUnknown_nonDelegating *PIUnknown_nonDelegating;

class SubComponentBase;
typedef SubComponentBase  *PSubComponentBase;
interface SubComponentBaseDelegate : public IUnknown, public IUnknown_nonDelegating
{
    /// Called when SubComponent is initialized to be fully usable (after first AddRef()).
    virtual void OnSubComponentInitialized(PSubComponentBase subComponent) = 0;
    /// Called when SubComponent is uninitialized to the no-more-usable state (after last Release()).
    virtual void OnSubComponentReleased(PSubComponentBase subComponent) = 0;
};
typedef SubComponentBaseDelegate  *PSubComponentBaseDelegate;

#define SUB_COMPONENT_IMPLEMENT_IUNKNOWN_AND_HELPERS(ClassName, COMInterfaceName)                           \
    HRESULT ClassName##::QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)     \
    {                                                                                                       \
        Utils::trace(_TEXT(__FUNCTION__));                                                                  \
        return COMLibrary::SubComponentBase::QueryInterface(riid, ppvObject);                               \
    }                                                                                                       \
                                                                                                            \
    ULONG STDMETHODCALLTYPE ClassName##::AddRef(void)                                                       \
    {                                                                                                       \
        Utils::trace(_TEXT(__FUNCTION__));                                                                  \
        return COMLibrary::SubComponentBase::AddRef();                                                      \
    }                                                                                                       \
                                                                                                            \
    ULONG STDMETHODCALLTYPE ClassName##::Release(void)                                                      \
    {                                                                                                       \
        Utils::trace(_TEXT(__FUNCTION__));                                                                  \
        return COMLibrary::SubComponentBase::Release();                                                     \
    }                                                                                                       \
                                                                                                            \
    PVOID ClassName##::TypecastSubComponentToCOMInterface()                                                 \
    {                                                                                                       \
        return static_cast<PVOID>(dynamic_cast<COMInterfaceName##*>(this));                                 \
    }

#define SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()                                                                    \
    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);   \
    virtual ULONG STDMETHODCALLTYPE AddRef(void);                                                                       \
    virtual ULONG STDMETHODCALLTYPE Release(void);                                                                      \
                                                                                                                        \
    virtual PVOID TypecastSubComponentToCOMInterface();


class SubComponentBase : public IUnknown, public IUnknown_nonDelegating
{
    LONG                        refCount{0};
    bool                        bInitialized{false};
    CRITICAL_SECTION            csInitialization;
    IID                         iid;
    bool                        bIgnore_AddRef_Release_Calls {false};

protected:
    PSubComponentBaseDelegate   delegateOfContainingCOMComponent;

public:
    SubComponentBase(REFIID riid);
    virtual ~SubComponentBase();

protected:
    virtual PVOID TypecastSubComponentToCOMInterface() = 0;

    // Subclass only has to override and call these functions (you can use macros SUB_COMPONENT_IMPLEMENT_IUNKNOWN()).
    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
    virtual ULONG STDMETHODCALLTYPE AddRef(void);
    virtual ULONG STDMETHODCALLTYPE Release(void);

    // Subclass doesn't have neither to override/implement, nor to call these functions.
    virtual HRESULT STDMETHODCALLTYPE QueryInterface_nonDelegating(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
    virtual ULONG STDMETHODCALLTYPE AddRef_nonDelegating(void);
    virtual ULONG STDMETHODCALLTYPE Release_nonDelegating(void);

    // Notifications to subclass. Called in thread-safe manner (synchronized).
    /// Called at first AddRef(). This is the final initialization point - SubComponent must be fully usable after this call.
    virtual void OnFirstAddRefInitialized();
    /// Called at last Release(). This is the first uninitialization point - SubComponent must not be usable after this call.
    virtual void OnLastReleaseUninitialized();

public:
    // These functions do nothing but may be overridden. Use them for (un)initializing included COM components.
    virtual HRESULT PostConstructorInitialize();
    virtual HRESULT PreDestructorRelease();

    REFIID GetIID();

private:
    void InitializeSubComponent();
    void ReleaseSubComponent();

    bool DecreaseRefCounterTo1IfNot0();

    friend class COMComponentBase;
};
typedef std::shared_ptr<SubComponentBase> SubComponentBasePtr;

///////////////////////////////////////////////////////////////////////////////////////////////////////

typedef std::pair<GUID, SubComponentBasePtr> PAIR_GUID_SUBCOMPONENT;
typedef std::pair<GUID, IUnknown*> PAIR_GUID_IUNKNOWN;
typedef std::map<PAIR_GUID_SUBCOMPONENT::first_type, PAIR_GUID_SUBCOMPONENT::second_type> DICT_SUBCOMPONENTS;
typedef std::map<PAIR_GUID_IUNKNOWN::first_type, PAIR_GUID_IUNKNOWN::second_type> DICT_AGGREGATED_INTERFACES;

class COMComponentBase : public SubComponentBaseDelegate
{
    bool                bInitialized{ false };
    CRITICAL_SECTION    csInitialization;
    LONG                cSubComponents{0};

    CLSID               clsid;

    IUnknown            *iUnknownOuter;
    PLONG               pcntrComponentsInServer;
    DICT_SUBCOMPONENTS  subcomponents;

    bool                bPostConstructorInitDone{false};
    
    // {CE2DBFC8-B1AF-4E99-9934-7D143863DD72}
    static const IID    UNINITIALIZED_PRIMARY_IID;
    IID                 primaryIID;

    bool                bAggregatedComponentIsAlsoAggregating{false};

protected:
    IUnknown                    *iUnknownInner{nullptr};
    DICT_AGGREGATED_INTERFACES  iUnknownInnerAggregatedInterfaces;

public:
    COMComponentBase(REFCLSID   clsid,
                     PLONG      pcntrComponentsInServer = nullptr,   ///< Address of active components counter within the COM server (DLL or EXE).
                     IUnknown*  iUnknownOuter = nullptr);
    virtual ~COMComponentBase();

protected:
    // Implementation of SubComponentBaseDelegate interface.
    virtual HRESULT STDMETHODCALLTYPE QueryInterface_nonDelegating(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
    virtual ULONG STDMETHODCALLTYPE AddRef_nonDelegating(void);
    virtual ULONG STDMETHODCALLTYPE Release_nonDelegating(void);

    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);
    virtual ULONG STDMETHODCALLTYPE AddRef(void);
    virtual ULONG STDMETHODCALLTYPE Release(void);

private:
    virtual void OnSubComponentInitialized(PSubComponentBase subComponent);
    virtual void OnSubComponentReleased(PSubComponentBase subComponent);

public:
    // Notifications to subclass. Called in thread-safe manner (synchronized).
    /// Called right before the first call to OnFirstAddRefInitialize() of the first subcomponent.
    /// This means, that no subcomponents are initialized yet (by the time of this call).
    /// For each SubComponent it's corresponding OnFirstAddRefInitialize() will be called later, when AddRef() of SubComponent is called.
    /// Use this function to prepare COM component in general.
    virtual void OnComponentInitialized(PSubComponentBase subComponent      ///< Pointer to the subcomponent, which will be initialized as first.
                                       );
    /// Called right after the last call to OnLastReleaseUninitialize() of the last subcomponent.
    /// This means, that all subcomponents are uninitialized already (by the time of this call).
    /// For each SubComponent it's corresponding OnLastReleaseUninitialize() has already been called before, when Release() of SubComponent was called.
    /// Use this function to release resources, posessed by the COM component.
    virtual void OnComponentUninitialized(PSubComponentBase subComponent    ///< Pointer to the subcomponent, which was uninitialized as last.
                                         );

    void AddSubComponent(SubComponentBasePtr subComponent, bool bPrimarySubcomponent = false);
    HRESULT AddAggregatedComponent(REFCLSID aggregatedCLSID, const std::vector<IID>& aggregatedIIDs, DWORD dwClsContext = CLSCTX_INPROC_SERVER);

public:
    PSubComponentBase GetPrimarySubcomponent();
    REFCLSID GetCLSID();

public:
    // These functions call corresponding functions of subcomponents - so call them if you override.
    // Use them to initialize included or aggregated COM components.
    // Note, that function COMComponentBase::PostConstructorInitialize() must be called after successful call to COMComponentBase::AddAggregatedComponent().
    virtual HRESULT PostConstructorInitialize();
    virtual HRESULT PreDestructorRelease();
};
typedef COMComponentBase *PCOMComponentBase;

}
