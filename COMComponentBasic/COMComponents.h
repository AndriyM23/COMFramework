#pragma once

#include "IComPtr.h"
#include "COM Component Kit\\COMComponentKit.h"
#include "Delivery Header Files\\Interfaces.h"

class SubComponentInterf01 : public COMLibrary::SubComponentBase, public IInterfBasic01
{
public:
    SubComponentInterf01();

    //Implementation of IInterfBasic01
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic01();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic01();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
};

class SubComponentInterf01Proxy : public COMLibrary::SubComponentBase, public IInterfBasic01
{
    COMLibrary::IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic01)>    includedInterf01;  // Included component.

public:
    SubComponentInterf01Proxy();

    //Proxying the implementation of included IInterfBasic04
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic01();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic01();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()

    virtual HRESULT PostConstructorInitialize();

#if 0
    // The smart pointer will be resetted in any case during deletion. So, there is no sense here to overload this function.
    virtual HRESULT PreDestructorRelease();
#endif
};

class SubComponentInterf02 : public COMLibrary::SubComponentBase, public IInterfBasic02
{
public:
    SubComponentInterf02();

    //Implementation of IInterfBasic02
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic02();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic02();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
};

class SubComponentInterf03 : public COMLibrary::SubComponentBase, public IInterfBasic03
{
public:
    SubComponentInterf03();

    //Implementation of IInterfBasic03
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic03();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic03();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
};

class SubComponentInterf04 : public COMLibrary::SubComponentBase, public IInterfBasic04
{
public:
    SubComponentInterf04();

    //Implementation of IInterfBasic04
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic04();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic04();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
};

class SubComponentInterf04Proxy : public COMLibrary::SubComponentBase, public IInterfBasic04
{
    COMLibrary::IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic04)>    includedInterf04;  // Included component.

public:
    SubComponentInterf04Proxy();

    //Proxying the implementation of included IInterfBasic04
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic04();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic04();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()

    virtual HRESULT PostConstructorInitialize();

    #if 0
    // The smart pointer will be resetted in any case during deletion. So, there is no sense here to overload this function.
    virtual HRESULT PreDestructorRelease();
    #endif
};

class SubComponentInterf05 : public COMLibrary::SubComponentBase, public IInterfBasic05
{
public:
    SubComponentInterf05();

    //Implementation of IInterfBasic05
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic05();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic05();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
};

class SubComponentInterf06 : public COMLibrary::SubComponentBase, public IInterfBasic06
{
public:
    SubComponentInterf06();

    //Implementation of IInterfBasic06
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic06();
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic06();

    // This subclass only has to override these functions and call those of the base class.
    SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
};


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class COMComponent01 : public COMLibrary::COMComponentBase
{
public:
    COMComponent01(PLONG        pcntrComponentsInServer,   ///< Address of active components counter within the COM server (DLL or EXE).
                   IUnknown*    iUnknownOuter = nullptr);
};
typedef COMComponent01 *PCOMComponent01;

class COMComponent04 : public COMLibrary::COMComponentBase
{
public:
    COMComponent04(PLONG        pcntrComponentsInServer,   ///< Address of active components counter within the COM server (DLL or EXE).
                   IUnknown*    iUnknownOuter = nullptr);
};
typedef COMComponent04 *PCOMComponent04;

class COMComponentI040506 : public COMLibrary::COMComponentBase
{
public:
    COMComponentI040506(PLONG        pcntrComponentsInServer,   ///< Address of active components counter within the COM server (DLL or EXE).
                          IUnknown*    iUnknownOuter = nullptr);
};
typedef COMComponentI040506 *PCOMComponentI040506;

class COMComponentI010203A04A05A06 : public COMLibrary::COMComponentBase
{
public:
    COMComponentI010203A04A05A06(PLONG  pcntrComponentsInServer    ///< Address of active components counter within the COM server (DLL or EXE).
                                );

    virtual HRESULT PostConstructorInitialize();
};
typedef COMComponentI010203A04A05A06 *PCOMComponentI010203A04A05A06;

class COMComponent02A01 : public COMLibrary::COMComponentBase
{
public:
    COMComponent02A01(PLONG  pcntrComponentsInServer,    ///< Address of active components counter within the COM server (DLL or EXE).
                      IUnknown*    iUnknownOuter = nullptr);

    virtual HRESULT PostConstructorInitialize();
};
typedef COMComponent02A01 *PCOMComponent02A01;

class COMComponentI04A02AA01 : public COMLibrary::COMComponentBase
{
public:
    COMComponentI04A02AA01(PLONG  pcntrComponentsInServer    ///< Address of active components counter within the COM server (DLL or EXE).
                          );

    virtual HRESULT PostConstructorInitialize();
};
typedef COMComponentI04A02AA01 *PCOMComponentI04A02AA01;