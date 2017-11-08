#pragma once

#include "COMComponentKit.h"

namespace COMLibrary
{
    class   COMComponentFactoryBase;
    typedef COMComponentFactoryBase*    PCOMComponentFactoryBase;
    typedef HRESULT  FUNCTION_CREATE_COMPONENT (IUnknown*                   pIUnknownOuter,
                                                REFIID                      riid,
                                                void**                      ppvObject,
                                                PCOMComponentFactoryBase    pComponentFactory);
    typedef FUNCTION_CREATE_COMPONENT*  PFUNCTION_CREATE_COMPONENT;

    class SubComponentFactoryBase : public COMLibrary::SubComponentBase, public IClassFactory
    {
        CLSID                       clsid;
        PFUNCTION_CREATE_COMPONENT  fnCreateComponent;
        PLONG                       pcntrServerLocks;

    public:
        SubComponentFactoryBase(REFCLSID                                clsid,
                                PFUNCTION_CREATE_COMPONENT              fnCreateComponent,
                                PLONG                                   pcntrServerLocks );

        // Implementation of IClassFactory.
        virtual HRESULT     STDMETHODCALLTYPE   CreateInstance(IUnknown*     pIUnknownOuter,
                                                               REFIID        riid,
                                                               void**        ppvObject);
        virtual HRESULT     STDMETHODCALLTYPE   LockServer(BOOL          fLock);

        // This subclass only has to override these functions and call those of the base class.
        SUB_COMPONENT_DECLARE_IUNKNOWN_AND_HELPERS()
    };
    typedef SubComponentFactoryBase*    PSubComponentFactoryBase;

    class COMComponentFactoryBase : public COMLibrary::COMComponentBase
    {
    public:
        /// One COM facrtory instance creates component of only one CLSID.
        COMComponentFactoryBase(REFCLSID                    clsid,
                                PFUNCTION_CREATE_COMPONENT  fnCreateComponent,
                                PLONG                       pcntrServerLocks );
    };
    typedef COMComponentFactoryBase*    PCOMComponentFactoryBase;
}
