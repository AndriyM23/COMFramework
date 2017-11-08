#pragma once

#include "IComPtr.h"

namespace COMLibrary
{
    // Here and further inspired by Kenny Kerr (https://msdn.microsoft.com/en-us/magazine/dn904668.aspx) and Dale Rogerson ("Inside COM").

    // DEFAULT CONSTRUCTOR
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom,riid_t>::IComPtr(REFIID riid) noexcept
        : iid(riid)
    {
    }

    // CONSTRUCTORs
#if 0
    // Commented out to enable to parameterize template as follows: IComPtr<IUnknown, IID_IUnknown>.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::IComPtr(const ICom* comInterface, REFIID riid) noexcept
        : comInterface(comInterface),
          iid(riid)
    {
        AddRef_internal();
    }
#endif

    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::IComPtr(const IUnknown* iUnknown, REFIID riid)
        : iid(riid)
    {
        if (iUnknown)
        {
            lastHres = iUnknown->QueryInterface(iid, reinterpret_cast<void**>(&comInterface));
            if (FAILED(lastHres))
            {
                comInterface = nullptr;
                if (bThrowsWinAPIException)
                    throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT(" Failed iUnknown(=%p)->QueryInterface(%s, &comInterface(=nullptr))"), iUnknown, GuidToString(iid).c_str());
            }
        }
    }

    // COPY CONSTRUCTOR
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::IComPtr(const IComPtr& source) noexcept
        : comInterface(source.comInterface),
          iid(source.iid),
          lastHres(source.lastHres)
    {
        AddRef_internal();
    }

    // COPY-LIKE-CONSTRUCTOR (but for differently parameterised templates).
    template<typename ICom, REFIID riid_t>
    template <typename ICom_source, REFIID riid_t_source>
    IComPtr<ICom, riid_t>::IComPtr(const IComPtr<ICom_source, riid_t_source>& source)
    {
        lastHres = source.Get()->QueryInterface(iid, reinterpret_cast<void**>(&comInterface));
        if (FAILED(lastHres))
        {
            comInterface = nullptr;
            if (bThrowsWinAPIException)
                throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT("(COPY-LIKE-CONSTRUCTOR) Failed source.Get()(=%p)->QueryInterface(%s, &comInterface(=nullptr))"), source.Get(), GuidToString(iid).c_str());
        }
    }

    // MOVE CONSTRUCTOR
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::IComPtr(IComPtr&& source) noexcept
        : comInterface(source.comInterface),
          iid(source.iid),
          lastHres(source.lastHres)
    {
        source.ReplaceWithNewIComPointer();  // Leaving the source in a "moved-from" state.
    }

    // MOVE-LIKE-CONSTRUCTOR (but for differently parameterised templates)
    template<typename ICom, REFIID riid_t>
    template <typename ICom_source, REFIID riid_t_source>
    IComPtr<ICom, riid_t>::IComPtr(IComPtr<ICom_source, riid_t_source>&& source)
    {
        lastHres = source.Get()->QueryInterface(iid, reinterpret_cast<void**>(&comInterface));
        if (FAILED(lastHres))
        {
            // Move attempt failed. Source is left unaltered. Destination(this object) is in "zero"-state.
            comInterface = nullptr;
            if (bThrowsWinAPIException)
                throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT("(MOVE-LIKE-CONSTRUCTOR) Failed source.Get()(=%p)->QueryInterface(%s, &comInterface(=nullptr))"), source.Get(), GuidToString(iid).c_str());
        }
        else
            // We don't have access to private method of another IComPtr template. So we use public method.
            source.Reset();  // Leaving the source in a "moved-from" state.
    }

    // DESTRUCTOR
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::~IComPtr()
    {
        ReplaceWithNewIComPointer();
    }

    // COPY ASSIGNMENT
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::operator=(const IComPtr& source) noexcept
    {
        return Copy(source);
    }

    // COPY-LIKE-ASSIGNMENT (but for differently parameterised templates).
    template<typename ICom, REFIID riid_t>
    template <typename ICom_source, REFIID riid_t_source>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::operator=(const IComPtr<ICom_source, riid_t_source>& source)
    {
        ICom* comInterface_local = nullptr;
        lastHres = source.Get()->QueryInterface(iid, reinterpret_cast<void**>(&comInterface_local));
        if (FAILED(lastHres))
        {
            comInterface_local = nullptr;
            if (bThrowsWinAPIException)
            {
                comInterface_local = comInterface;
                ReplaceWithNewIComPointer();
                throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT("(COPY-LIKE-ASSIGNMENT) Failed source.Get()(=%p)->QueryInterface(%s, &comInterface_local(=nullptr)), comInterface was %p"), source.Get(), GuidToString(iid).c_str(), comInterface_local);
            }
        }
        ReplaceWithNewIComPointer(comInterface_local, false);
        return *this;
    }

    // MOVE ASSIGNMENT
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::operator=(IComPtr&& source) noexcept
    {
        if (&source != this)
        {
            ReplaceWithNewIComPointer(source);
            source.ReplaceWithNewIComPointer();  // Leaving the source in a "moved-from" state.
        }
        return *this;
    }

    // MOVE-LIKE-ASSIGNMENT (but for differently parameterised templates).
    template<typename ICom, REFIID riid_t>
    template <typename ICom_source, REFIID riid_t_source>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::operator=(IComPtr<ICom_source, riid_t_source>&& source)
    {
        ICom* comInterface_local = nullptr;
        lastHres = source.Get()->QueryInterface(iid, reinterpret_cast<void**>(&comInterface_local));
        if (FAILED(lastHres))
        {
            // Move attempt failed. Source is left unaltered. Destination(this object) is in "zero"-state.
            comInterface_local = comInterface; // This is for exception only.
            ReplaceWithNewIComPointer();
            if (bThrowsWinAPIException)
                throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT("(MOVE-LIKE-ASSIGNMENT) Failed source.Get()(=%p)->QueryInterface(%s, &comInterface_local(=nullptr)), comInterface was %p"), source.Get(), GuidToString(iid).c_str(), comInterface_local);
        }
        else
        {
            // Move attempt succeeded. Source is left in "moved-from"-state. Destination(this object) is now as source.
            ReplaceWithNewIComPointer(comInterface_local, false);
            // We don't have access to private method of another IComPtr template. So we use public method.
            source.Reset();  // Leaving the source in a "moved-from" state.
        }
        return *this;
    }

    // ASSIGNMENTs
#if 0
    // Commented out to enable to parameterize template as follows: IComPtr<IUnknown, IID_IUnknown>.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::operator=(const ICom* source) noexcept
    {
        return Copy(source);
    }
#endif

    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::operator=(const IUnknown* iUnknown) noexcept
    {
        if (iUnknown == nullptr)
            ReplaceWithNewIComPointer();
        else
        {
            ICom* comInterface_local = nullptr;
            lastHres = iUnknown->QueryInterface(iid, reinterpret_cast<void**>(&comInterface_local));
            if (FAILED(lastHres))
            {
                comInterface_local = comInterface; // This is for exception only.
                ReplaceWithNewIComPointer();
                if (bThrowsWinAPIException)
                    throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT(" Failed iUnknown(=%p)->QueryInterface(%s, &comInterface_local(=nullptr)), comInterface was %p"), iUnknown, GuidToString(iid).c_str(), comInterface_local);
            }
            else
                ReplaceWithNewIComPointer(comInterface_local, false);
        }
        return *this;
    }

    /// Typecast operator to ICom*.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::operator ICom*() const noexcept
    {
        return comInterface;
    }

    /// May be needed, if necessary to fill in the internal ICom pointer from outside, e.g. such as in a call to QueryInterface().
    template<typename ICom, REFIID riid_t>
    ICom** IComPtr<ICom, riid_t>::operator &()
    {
        return GetIComPointerAddress();
    }

#if 0
    ICom& operator *() ... //It is better not to expose this functionality. What should be done with dereferenced pointer to a COM Interface ??? It has no sense.
#endif

    /// \note Here we prohibit possibility to call AddRef()&Release() with this pointer (by Kenny Kerr: https://msdn.microsoft.com/en-us/magazine/dn904668.aspx)
    template<typename ICom, REFIID riid_t>
    IComNoAddRefRelease<ICom>* IComPtr<ICom, riid_t>::operator ->()
    {
        if (comInterface == nullptr)
            if (bThrowsWinAPIException)
            {
                lastHres = E_POINTER;
                throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT(" Internal COM interface pointer == nullptr"));
            }
            else
            {
                _wassert(TEXT_TO_UNICODE(__FUNCTION__) L" Internal COM interface pointer == nullptr", TEXT_TO_UNICODE(__FILE__), __LINE__);
            }
        return static_cast<IComNoAddRefRelease<ICom> * >(comInterface);
    }

    template<typename ICom, REFIID riid_t>
    bool IComPtr<ICom, riid_t>::operator!() const noexcept
    {
        return (comInterface == nullptr);
    }

    /// Typecast operator to bool.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>::operator bool() const noexcept
    {
        return (comInterface != nullptr);
    }

    // METAPROPERTIES
    template<typename ICom, REFIID riid_t>
    REFIID IComPtr<ICom, riid_t>::GetIID() const noexcept
    {
        return iid;
    }

    template<typename ICom, REFIID riid_t>
    bool& IComPtr<ICom, riid_t>::Throws() noexcept
    {
        return bThrowsWinAPIException;
    }

    template<typename ICom, REFIID riid_t>
    HRESULT IComPtr<ICom, riid_t>::GetLastHRESULT() const noexcept
    {
        return lastHres;
    }

    /// \note This class may misbehave, if, using the returned pointer, COM interface reference counter is manipulated outside IComPtr (damage of invariant).
    template<typename ICom, REFIID riid_t>
    ICom* IComPtr<ICom, riid_t>::Get() const noexcept
    {
        return comInterface;
    }

    // MANIPULATION METHODS
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Reset(const ICom* other) noexcept
    {
        ReplaceWithNewIComPointer(other);
        return *this;
    }

    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Reset(const IComPtr& other) noexcept
    {
        ReplaceWithNewIComPointer(other);
        return *this;
    }

    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Swap(IComPtr& other) noexcept
    {
        std::swap(this, other);
        return *this;
    }

    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Copy(const ICom* source) noexcept
    {
        ReplaceWithNewIComPointer(source);
        return *this;
    }

    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Copy(const IComPtr& source) noexcept
    {
        ReplaceWithNewIComPointer(source);
        return *this;
    }

    /// \note This function does not increment interface reference counter. Use with care.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Attach(const ICom* other) noexcept
    {
        ReplaceWithNewIComPointer();
        comInterface = other;
        return *this;
    }

    /// \note This function does not increment interface reference counter. Use with care.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::Attach(const IComPtr& other) noexcept
    {
        ReplaceWithNewIComPointer();
        comInterface = other.comInterface;
        iid = other.iid;
        lastHres = other.lastHres;
        return *this;
    }

    /// \note This function does not decrement interface reference counter. Use with care.
    template<typename ICom, REFIID riid_t>
    ICom* IComPtr<ICom, riid_t>::Detach() noexcept
    {
        ICom* comInterface_retVal = comInterface;
        comInterface = nullptr;
        return comInterface_retVal;
    }

    /// May be needed, if necessary to fill in the internal ICom pointer from outside, e.g. such as in a call to QueryInterface().
    template<typename ICom, REFIID riid_t>
    ICom** IComPtr<ICom, riid_t>::GetIComPointerAddress()
    {
        if (comInterface != nullptr)
            if (bThrowsWinAPIException)
            {
                lastHres = E_POINTER;
                throw WinAPIException(lastHres, _TEXT(__FUNCTION__) _TEXT(" Internal COM interface pointer(%p) != nullptr - overwiting it will damage the invariant of the class IComPtr."), comInterface);
            }
            else
            {
                _wassert(TEXT_TO_UNICODE(__FUNCTION__) L" Internal COM interface pointer != nullptr - overwiting it will damage the invariant of the class IComPtr.", TEXT_TO_UNICODE(__FILE__), __LINE__);
            }
        return &comInterface;
    }

    /// Helps to avoid call to QueryInterface(), if duplicated pointer is to the same COM interface, as this one.
    template<typename ICom, REFIID riid_t>
    IComPtr<ICom, riid_t>& IComPtr<ICom, riid_t>::DuplicateIComPointer(ICom** pDuplicatedIComPointer) noexcept
    {
        AddRef_internal();
        *pDuplicatedIComPointer = comInterface;
        return *this;
    }

    /// We can obtain IComPtr to another COM interface.
    /// We can use this IComPtr further, or just once to see, whether the underlying COM component supports the requested COM interface.
    template<typename ICom, REFIID riid_t>
    template <typename ICom_another, REFIID riid_t_another>
    IComPtr<ICom_another, riid_t_another> IComPtr<ICom, riid_t>::GetAnotherIComPtr() const
    {
        IComPtr<ICom_another, riid_t_another> comInterface_another;
        lastHres = comInterface->QueryInterface(riid_t_another, reinterpret_cast<void**>(comInterface_another.GetIComPointerAddress()));
        return comInterface_another;
    }

    template<typename ICom, REFIID riid_t>
    _Check_return_ HRESULT IComPtr<ICom, riid_t>::CoCreateInstance( _In_     REFCLSID    rclsid,
                                                                    _In_opt_ LPUNKNOWN   pUnkOuter,
                                                                    _In_     DWORD       dwClsContext )
    {
        return (lastHres = ::CoCreateInstance(rclsid,
                                              pUnkOuter,
                                              dwClsContext,
                                              iid,
                                              reinterpret_cast<LPVOID*>(GetIComPointerAddress())));
    }

    // INTERNAL SERVICE METHODS
    template<typename ICom, REFIID riid_t>
    ULONG IComPtr<ICom, riid_t>::AddRef_internal() noexcept
    {
        ULONG retVal = 0;
        if (comInterface)
            retVal = comInterface->AddRef();
        return retVal;
    }

    template<typename ICom, REFIID riid_t>
    ULONG IComPtr<ICom, riid_t>::ReplaceWithNewIComPointer(const ICom* newComInterface, bool doAddRef) noexcept
    {
        ULONG retVal = 0;

        if (newComInterface == nullptr)
        {
            if (comInterface)
            {
                // Temp variable protects us from possible iteration into this function again (from within comInterface_old->Release() call).
                ICom* comInterface_old = comInterface;
                comInterface = nullptr;
                retVal = comInterface_old->Release();
            }
        }
        else if (newComInterface != comInterface)
        {
            // Temp variable guards us from possible early-deletion of the object, in case if newComInterface == comInterface.
            // Although, this situation is already checked in IF-statement, it is good to write in this style.
            ICom* comInterface_old = comInterface;
            comInterface = const_cast<ICom*>(newComInterface);
            if (doAddRef)
                retVal = AddRef_internal();
            if (comInterface_old)
                comInterface_old->Release();
        }

        return retVal;
    }

    template<typename ICom, REFIID riid_t>
    ULONG IComPtr<ICom, riid_t>::ReplaceWithNewIComPointer(const IComPtr& newComInterface) noexcept
    {
        if (&newComInterface != this)
        {
            ULONG retVal = ReplaceWithNewIComPointer(newComInterface.comInterface);
            iid = newComInterface.iid;
            lastHres = newComInterface.lastHres;
            return retVal;
        }
        else
            // We do nothing if we assign/replace us with ourself.
            return 0;
    }


    // Non-member IComPtr operations&operators.

    template <typename ICom>
    void swap(IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept
    {
        left.Swap(right);
    }

    template <typename ICom>
    bool operator == (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept
    {
        return (left.Get() == right.Get());
    }

    template <typename ICom>
    bool operator != (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept
    {
        return (left.Get() != right.Get());
    }

#if 0
    There is probably no sense in the following comparisons, because the pointers compared are pointers to COM interfaces.
    bool operator < (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept  ...
    bool operator > (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept  ...
    bool operator <= (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept ...
    bool operator >= (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept ...
#endif
}