#include "Utils.h"
#include "COMComponents.h"
#include "COM Component Kit\\COMComponentFactoryKit.h"
#include "COM Component Kit\\COMComponentRegistrator.h"

using namespace std;
using namespace Utils;
using namespace COMLibrary;

static HINSTANCE    hInstanceDll { nullptr };
static LONG         cntrServerLocks { 0 };
static LONG         cntrComponents { 0 };

////////////////////////////////////////////////////////////////////////////////////////////////////

FUNCTION_CREATE_COMPONENT CreateCOMComponent01;
FUNCTION_CREATE_COMPONENT CreateCOMComponent04;
FUNCTION_CREATE_COMPONENT CreateCOMComponentI040506;
FUNCTION_CREATE_COMPONENT CreateCOMComponentI010203A04A05A06;
FUNCTION_CREATE_COMPONENT CreateCOMComponent02A01;
FUNCTION_CREATE_COMPONENT CreateCOMComponentI04A02AA01;

typedef pair<CLSID, PFUNCTION_CREATE_COMPONENT> CLSID_FUNCCREATCMPNT_PAIR;
static CLSID_FUNCCREATCMPNT_PAIR functionsCreateComponentArr[] =
{
    { CLSID_COMComponent01,                 CreateCOMComponent01 },
    { CLSID_COMComponent04,                 CreateCOMComponent04 },
    { CLSID_COMComponentI040506,            CreateCOMComponentI040506 },
    { CLSID_COMComponentI010203A04A05A06,   CreateCOMComponentI010203A04A05A06 },
    { CLSID_COMComponent02A01,              CreateCOMComponent02A01 },
    { CLSID_COMComponentI04A02AA01,         CreateCOMComponentI04A02AA01 }
};

typedef map<CLSID_FUNCCREATCMPNT_PAIR::first_type, CLSID_FUNCCREATCMPNT_PAIR::second_type> DICTTYPE_CREATION_FUNCTIONS;
static DICTTYPE_CREATION_FUNCTIONS creationFunctionsDict {functionsCreateComponentArr, functionsCreateComponentArr + ELEMENTS_IN_ARRAY(functionsCreateComponentArr)};

HRESULT CreateCOMComponent01(IUnknown*                  pIUnknownOuter,
                             REFIID                     riid,
                             void**                     ppvObject,
                             PCOMComponentFactoryBase   pComponentFactory)
{
    UNREFERENCED_PARAMETER(pComponentFactory);

    assert(CLSID_COMComponent01 == pComponentFactory->GetCLSID());

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|pUnknownOuter=%p, riid=%s, clsid=%s"),
        pIUnknownOuter,
        GuidToString(riid).c_str(),
        GuidToString(pComponentFactory->GetCLSID()).c_str());

    if (pIUnknownOuter && riid != IID_IUnknown)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_NOAGGREGATION"));
        return CLASS_E_NOAGGREGATION;
    }

    PCOMComponent01 pComponent = new COMComponent01(&cntrComponents, pIUnknownOuter);
    if (!pComponent)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    IUnknown* primarySubComponent_IUnknown;
    if (pIUnknownOuter)
    {
        // We are being aggregated.
        IUnknown_nonDelegating* primarySubComponentNonDeleg = static_cast<IUnknown_nonDelegating*>(pComponent->GetPrimarySubcomponent());
        primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(primarySubComponentNonDeleg);
    }
    else
        // We are not being aggregated.
        primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(pComponent->GetPrimarySubcomponent());

    HRESULT result = primarySubComponent_IUnknown->QueryInterface(riid, ppvObject);
    if (FAILED(result))
    {
        trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface_nonDelegating(%s) failed. Result=%08.8Xh. Delete subcomponent."), GuidToString(riid).c_str(), result);
        *ppvObject = nullptr;
        delete pComponent;
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh, *ppvObject=%p"), result, *ppvObject);
    return result;
}

HRESULT CreateCOMComponent04(IUnknown*                  pIUnknownOuter,
                             REFIID                     riid,
                             void**                     ppvObject,
                             PCOMComponentFactoryBase   pComponentFactory)
{
    UNREFERENCED_PARAMETER(pComponentFactory);

    assert(CLSID_COMComponent04 == pComponentFactory->GetCLSID());

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|pUnknownOuter=%p, riid=%s, clsid=%s"),
        pIUnknownOuter,
        GuidToString(riid).c_str(),
        GuidToString(pComponentFactory->GetCLSID()).c_str());

    if (pIUnknownOuter && riid != IID_IUnknown)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_NOAGGREGATION"));
        return CLASS_E_NOAGGREGATION;
    }

    PCOMComponent04 pComponent = new COMComponent04(&cntrComponents, pIUnknownOuter);
    if (!pComponent)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    IUnknown* primarySubComponent_IUnknown;
    if (pIUnknownOuter)
    {
        // We are being aggregated.
        IUnknown_nonDelegating* primarySubComponentNonDeleg = static_cast<IUnknown_nonDelegating*>(pComponent->GetPrimarySubcomponent());
        primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(primarySubComponentNonDeleg);
    }
    else
        // We are not being aggregated.
        primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(pComponent->GetPrimarySubcomponent());

    HRESULT result = primarySubComponent_IUnknown->QueryInterface(riid, ppvObject);
    if (FAILED(result))
    {
        trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface_nonDelegating(%s) failed. Result=%08.8Xh. Delete subcomponent."), GuidToString(riid).c_str(), result);
        *ppvObject = nullptr;
        delete pComponent;
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh, *ppvObject=%p"), result, *ppvObject);
    return result;
}

HRESULT CreateCOMComponentI040506(IUnknown*                  pIUnknownOuter,
                                    REFIID                     riid,
                                    void**                     ppvObject,
                                    PCOMComponentFactoryBase   pComponentFactory)
{
    UNREFERENCED_PARAMETER(pComponentFactory);

    assert(CLSID_COMComponentI040506 == pComponentFactory->GetCLSID());

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|pUnknownOuter=%p, riid=%s, clsid=%s"),
        pIUnknownOuter,
        GuidToString(riid).c_str(),
        GuidToString(pComponentFactory->GetCLSID()).c_str());

    if (pIUnknownOuter && riid != IID_IUnknown)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_NOAGGREGATION"));
        return CLASS_E_NOAGGREGATION;
    }

    PCOMComponentI040506 pComponent = new COMComponentI040506(&cntrComponents, pIUnknownOuter);
    if (!pComponent)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    HRESULT result = pComponent->PostConstructorInitialize();
    if (FAILED(result))
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|pComponent->PostConstructorInitialize() failed.|Result=%08.8Xh"), result);
        return result;
    }

    IUnknown* primarySubComponent_IUnknown;
    if (pIUnknownOuter)
    {
        // We are being aggregated.
        IUnknown_nonDelegating* primarySubComponentNonDeleg = static_cast<IUnknown_nonDelegating*>(pComponent->GetPrimarySubcomponent());
        primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(primarySubComponentNonDeleg);
    }
    else
        // We are not being aggregated.
        primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(pComponent->GetPrimarySubcomponent());

    result = primarySubComponent_IUnknown->QueryInterface(riid, ppvObject);
    if (FAILED(result))
    {
        trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface_nonDelegating(%s) failed. Result=%08.8Xh. Delete subcomponent."), GuidToString(riid).c_str(), result);
        *ppvObject = nullptr;
        delete pComponent;
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh, *ppvObject=%p"), result, *ppvObject);
    return result;
}

HRESULT CreateCOMComponentI010203A04A05A06(IUnknown*                  pIUnknownOuter,
                                           REFIID                     riid,
                                           void**                     ppvObject,
                                           PCOMComponentFactoryBase   pComponentFactory)
{
    UNREFERENCED_PARAMETER(pComponentFactory);

    assert(CLSID_COMComponentI010203A04A05A06 == pComponentFactory->GetCLSID());

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|pUnknownOuter=%p, riid=%s, clsid=%s"),
        pIUnknownOuter,
        GuidToString(riid).c_str(),
        GuidToString(pComponentFactory->GetCLSID()).c_str());

    if (pIUnknownOuter)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_NOAGGREGATION"));
        return CLASS_E_NOAGGREGATION;
    }

    PCOMComponentI010203A04A05A06 pComponent = new COMComponentI010203A04A05A06(&cntrComponents);
    if (!pComponent)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    // If PostConstructorInitialize() succeeds, we have one reference onto the created outer COM component.
    HRESULT result = pComponent->PostConstructorInitialize();
    if (FAILED(result))
    {
        *ppvObject = nullptr;
        delete pComponent;
        trace(_TEXT(__FUNCTION__) _TEXT("|pComponent->PostConstructorInitialize() failed.|Result=%08.8Xh"), result);
    }
    else
    {
        IUnknown* primarySubComponent = static_cast<IUnknown*>(pComponent->GetPrimarySubcomponent());
        // If QueryInterface() succeeds, we have 2-nd reference onto the created outer COM component.
        result = primarySubComponent->QueryInterface(riid, ppvObject);
        if (FAILED(result))
            trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface(%s) failed. Result=%08.8Xh. Delete subcomponent."), GuidToString(riid).c_str(), result);
        if (SUCCEEDED(result))
            trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface(%s) succeeded. Result=%08.8Xh."), GuidToString(riid).c_str(), result);
        // Deleting either one additional reference or deleting the created outer COM component.
        primarySubComponent->Release();
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh, *ppvObject=%p"), result, *ppvObject);
    return result;
}

HRESULT CreateCOMComponent02A01(IUnknown*                  pIUnknownOuter,
                                REFIID                     riid,
                                void**                     ppvObject,
                                PCOMComponentFactoryBase   pComponentFactory)
{
    UNREFERENCED_PARAMETER(pComponentFactory);

    assert(CLSID_COMComponent02A01 == pComponentFactory->GetCLSID());

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|pUnknownOuter=%p, riid=%s, clsid=%s"),
        pIUnknownOuter,
        GuidToString(riid).c_str(),
        GuidToString(pComponentFactory->GetCLSID()).c_str());

    if (pIUnknownOuter && riid != IID_IUnknown)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_NOAGGREGATION"));
        return CLASS_E_NOAGGREGATION;
    }

    PCOMComponent02A01 pComponent = new COMComponent02A01(&cntrComponents, pIUnknownOuter);
    if (!pComponent)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    // If PostConstructorInitialize() succeeds, we have one reference onto the created outer COM component.
    HRESULT result = pComponent->PostConstructorInitialize();
    if (FAILED(result))
    {
        *ppvObject = nullptr;
        delete pComponent;
        trace(_TEXT(__FUNCTION__) _TEXT("|pComponent->PostConstructorInitialize() failed.|Result=%08.8Xh"), result);
    }
    else
    {
        IUnknown* primarySubComponent_IUnknown;
        if (pIUnknownOuter)
        {
            // We are being aggregated.
            IUnknown_nonDelegating* primarySubComponentNonDeleg = static_cast<IUnknown_nonDelegating*>(pComponent->GetPrimarySubcomponent());
            primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(primarySubComponentNonDeleg);
        }
        else
            // We are not being aggregated.
            primarySubComponent_IUnknown = reinterpret_cast<IUnknown*>(pComponent->GetPrimarySubcomponent());

        // If QueryInterface() succeeds, we have 2-nd reference onto the created outer COM component.
        result = primarySubComponent_IUnknown->QueryInterface(riid, ppvObject);

        if (FAILED(result))
            trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface_nonDelegating(%s) failed. Result=%08.8Xh. Delete subcomponent."), GuidToString(riid).c_str(), result);
        if (SUCCEEDED(result))
            trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface_nonDelegating(%s) succeeded. Result=%08.8Xh."), GuidToString(riid).c_str(), result);

        // Deleting either one additional reference or deleting the created outer COM component.
        primarySubComponent_IUnknown->Release();
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh, *ppvObject=%p"), result, *ppvObject);
    return result;
}

HRESULT CreateCOMComponentI04A02AA01(IUnknown*                  pIUnknownOuter,
                                     REFIID                     riid,
                                     void**                     ppvObject,
                                     PCOMComponentFactoryBase   pComponentFactory)
{
    UNREFERENCED_PARAMETER(pComponentFactory);

    assert(CLSID_COMComponentI04A02AA01 == pComponentFactory->GetCLSID());

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|pUnknownOuter=%p, riid=%s, clsid=%s"),
        pIUnknownOuter,
        GuidToString(riid).c_str(),
        GuidToString(pComponentFactory->GetCLSID()).c_str());

    if (pIUnknownOuter)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_NOAGGREGATION"));
        return CLASS_E_NOAGGREGATION;
    }

    PCOMComponentI04A02AA01 pComponent = new COMComponentI04A02AA01(&cntrComponents);
    if (!pComponent)
    {
        *ppvObject = nullptr;
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    // If PostConstructorInitialize() succeeds, we have one reference onto the created outer COM component.
    HRESULT result = pComponent->PostConstructorInitialize();
    if (FAILED(result))
    {
        *ppvObject = nullptr;
        delete pComponent;
        trace(_TEXT(__FUNCTION__) _TEXT("|pComponent->PostConstructorInitialize() failed.|Result=%08.8Xh"), result);
    }
    else
    {
        IUnknown* primarySubComponent = static_cast<IUnknown*>(pComponent->GetPrimarySubcomponent());
        // If QueryInterface() succeeds, we have 2-nd reference onto the created outer COM component.
        result = primarySubComponent->QueryInterface(riid, ppvObject);
        if (FAILED(result))
            trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface(%s) failed. Result=%08.8Xh. Delete subcomponent."), GuidToString(riid).c_str(), result);
        if (SUCCEEDED(result))
            trace(_TEXT(__FUNCTION__) _TEXT("pComponent->QueryInterface(%s) succeeded. Result=%08.8Xh."), GuidToString(riid).c_str(), result);
        // Deleting either one additional reference or deleting the created outer COM component.
        primarySubComponent->Release();
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|result=%08.8Xh, *ppvObject=%p"), result, *ppvObject);
    return result;
}


////////////////////////////////////////////////////////////////////////////////////////////////////

static COMComponentsInfo comComponentsInThisDLLServer;
static void InitCOMComponentsInThisDLLServer()
{
    COMComponentInfoPtr comComponentInfo{ new COMComponentInfo{ hInstanceDll,
                                                                CLSID_COMComponent01,
                                                                _TEXT(COMCOMPONENT01_FRIENDLY_NAME),
                                                                _TEXT(COMCOMPONENT01_PROGID),
                                                                _TEXT(COMCOMPONENT01_VERINDEPPROGID),
                                                                { COMCategoryIFB01 },
                                                                {} } };
    comComponentsInThisDLLServer.push_back(comComponentInfo);

    comComponentInfo.reset(new COMComponentInfo{hInstanceDll,
                                                CLSID_COMComponent04,
                                                _TEXT(COMCOMPONENT04_FRIENDLY_NAME),
                                                _TEXT(COMCOMPONENT04_PROGID),
                                                _TEXT(COMCOMPONENT04_VERINDEPPROGID),
                                                { COMCategoryIFB04 },
                                                {} });
    comComponentsInThisDLLServer.push_back(comComponentInfo);

    comComponentInfo.reset(new COMComponentInfo{hInstanceDll,
                                                CLSID_COMComponentI040506,
                                                _TEXT(COMCOMPONENTINC040506_FRIENDLY_NAME),
                                                _TEXT(COMCOMPONENTINC040506_PROGID),
                                                _TEXT(COMCOMPONENTINC040506_VERINDEPPROGID),
                                                { COMCategoryIFB040506 },
                                                { COMCategoryIFB04 } });
    comComponentsInThisDLLServer.push_back(comComponentInfo);

    comComponentInfo.reset(new COMComponentInfo{hInstanceDll,
                                                CLSID_COMComponentI010203A04A05A06,
                                                _TEXT(COMCOMPONENTI010203A04A05A06_FRIENDLY_NAME),
                                                _TEXT(COMCOMPONENTI010203A04A05A06_PROGID),
                                                _TEXT(COMCOMPONENTI010203A04A05A06_VERINDEPPROGID),
                                                { COMCategoryIFB01, COMCategoryIFB04, COMCategoryIFB040506, COMCategoryIFB010203040506 },
                                                { COMCategoryIFB01, COMCategoryIFB040506 } });
    comComponentsInThisDLLServer.push_back(comComponentInfo);

    comComponentInfo.reset(new COMComponentInfo{hInstanceDll,
                                                CLSID_COMComponent02A01,
                                                _TEXT(COMCOMPONENT02A01_FRIENDLY_NAME),
                                                _TEXT(COMCOMPONENT02A01_PROGID),
                                                _TEXT(COMCOMPONENT02A01_VERINDEPPROGID),
                                                { COMCategoryIFB01 },
                                                { COMCategoryIFB01 } });
    comComponentsInThisDLLServer.push_back(comComponentInfo);

    comComponentInfo.reset(new COMComponentInfo{hInstanceDll,
                                                CLSID_COMComponentI04A02AA01,
                                                _TEXT(COMCOMPONENTI04A02AA01_FRIENDLY_NAME),
                                                _TEXT(COMCOMPONENTI04A02AA01_PROGID),
                                                _TEXT(COMCOMPONENTI04A02AA01_VERINDEPPROGID),
                                                { COMCategoryIFB01, COMCategoryIFB04 },
                                                { COMCategoryIFB01, COMCategoryIFB04 } });
    comComponentsInThisDLLServer.push_back(comComponentInfo);
}

BOOL WINAPI DllMain(HINSTANCE hInstDll, DWORD fdwReason, LPVOID lpvReserved)
{
    UNREFERENCED_PARAMETER(lpvReserved);

    switch (fdwReason)
    {
    case DLL_PROCESS_ATTACH:
        hInstanceDll = hInstDll;
        InitCOMComponentsInThisDLLServer();
        trace(_TEXT(__FUNCTION__) _TEXT("|<<<|DLL_PROCESS_ATTACH|hInstanceDll=%p"), hInstanceDll);
        break;

    case DLL_PROCESS_DETACH:
        trace(_TEXT(__FUNCTION__) _TEXT("|<<<|DLL_PROCESS_DETACH|hInstDll=%p, hInstanceDll=%p"), hInstDll, hInstanceDll);
        break;

    case DLL_THREAD_ATTACH:
        trace(_TEXT(__FUNCTION__) _TEXT("|<<<|DLL_THREAD_ATTACH|hInstDll=%p, hInstanceDll=%p"), hInstDll, hInstanceDll);
        break;

    case DLL_THREAD_DETACH:
        trace(_TEXT(__FUNCTION__) _TEXT("|<<<|DLL_THREAD_DETACH|hInstDll=%p, hInstanceDll=%p"), hInstDll, hInstanceDll);
        break;
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>"));
    return TRUE;
}

STDAPI DllRegisterServer()
try
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    COMComponentsInfo::iterator iComComponent = comComponentsInThisDLLServer.begin();
    while (iComComponent != comComponentsInThisDLLServer.end())
        (*iComComponent++)->Register();
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|hres=S_OK"));
    return S_OK;
}
catch (const WinAPIException& e)
{
    _tstring reason;
    e.what(reason);
    trace(_TEXT(__FUNCTION__) _TEXT("|Registration failed. Reason: %s"), reason.c_str());
    try
    {
        COMComponentsInfo::iterator iComComponent = comComponentsInThisDLLServer.begin();
        while (iComComponent != comComponentsInThisDLLServer.end())
            (*iComComponent++)->Unregister();
    }
    catch (const WinAPIException&)
    { }

    return E_UNEXPECTED;
}


STDAPI DllUnregisterServer()
try
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<"));
    COMComponentsInfo::iterator iComComponent = comComponentsInThisDLLServer.begin();
    while (iComComponent != comComponentsInThisDLLServer.end())
        (*iComComponent++)->Unregister();
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|hres=S_OK"));
    return S_OK;
}
catch (const WinAPIException& e)
{
    _tstring reason;
    e.what(reason);
    trace(_TEXT(__FUNCTION__) _TEXT("|Registration failed. Reason: %s"), reason.c_str());
    return E_UNEXPECTED;
}


STDAPI DllGetClassObject(REFCLSID   clsid,
                         REFIID     iid,
                         void**     ppv)
{
    if (creationFunctionsDict.find(clsid) == creationFunctionsDict.end())
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:CLASS_E_CLASSNOTAVAILABLE|clsid=%s"), GuidToString(clsid).c_str());
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|clsid=%s|iid=%s|creationFunction=%p"), GuidToString(clsid).c_str(),
                                                                                 GuidToString(iid).c_str(),
                                                                                 creationFunctionsDict[clsid]);

    PCOMComponentFactoryBase pComponentFactory = new COMComponentFactoryBase(clsid, creationFunctionsDict[clsid], &cntrServerLocks);
    if (!pComponentFactory)
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:E_OUTOFMEMORY"));
        return E_OUTOFMEMORY;
    }

    PSubComponentFactoryBase pSubComponentFactory = dynamic_cast<PSubComponentFactoryBase>(pComponentFactory->GetPrimarySubcomponent());
    HRESULT hResult = pSubComponentFactory->QueryInterface(iid, ppv);
    if (FAILED(hResult))
    {
        trace(_TEXT(__FUNCTION__) _TEXT("pSubComponentFactory->QueryInterface(%s) failed. result=%08.8Xh. Delete factory."), GuidToString(iid).c_str(), hResult);
        delete pComponentFactory;
    }

    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:%08.8Xh|ppv=%p|clsid=%s|iid=%s|creationFunction=%p"),
                    hResult,
                    *ppv,
                    GuidToString(clsid).c_str(),
                    GuidToString(iid).c_str(),
                    creationFunctionsDict[clsid]);
    return hResult;
}

STDAPI DllCanUnloadNow()
{
    trace(_TEXT(__FUNCTION__) _TEXT("|<<<|cntrComponents=%u, cntrServerLocks=%u"), cntrComponents, cntrServerLocks);
    if (cntrComponents == 0 && cntrServerLocks == 0)
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:S_OK"));
        return S_OK;
    }
    else
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|>>>|RET:S_FALSE"));
        return S_FALSE;
    }
}
