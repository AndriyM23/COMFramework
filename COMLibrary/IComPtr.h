#pragma once

#include "WinAPIException.h"

#include <algorithm>        // For std::swap().
#include <utility>          // For std::move() - my be used by class' users.
#include <assert.h>         // For _wassert().

namespace COMLibrary
{
    // Here and further inspired by Kenny Kerr (https://msdn.microsoft.com/en-us/magazine/dn904668.aspx) and Dale Rogerson (book "Inside COM").
    
    // Used in operator ->().
    template<typename ICom> 
    class IComNoAddRefRelease : public ICom
    {
        virtual ULONG STDMETHODCALLTYPE AddRef(void);
        virtual ULONG STDMETHODCALLTYPE Release(void);
    };

    #define ICOMPTR_TEMPLATE_ARGUMENTS(IComName) IComName, IID_##IComName
    template<typename ICom, REFIID riid_t = __uuidof(ICom)>
    class IComPtr
    {
        ICom*       comInterface {nullptr};
        IID         iid{ riid_t };

        static const bool bThrowsWinAPIException_DEFAULT{ true };
        bool        bThrowsWinAPIException {bThrowsWinAPIException_DEFAULT};
        HRESULT     lastHres{S_OK};

    public:
        // DEFAULT CONSTRUCTOR
        explicit IComPtr(REFIID riid = riid_t) noexcept;

        // CONSTRUCTORs
        #if 0
        // Commented out to enable to parameterize template as follows: IComPtr<IUnknown, IID_IUnknown>.
        explicit IComPtr(const ICom* comInterface, REFIID riid = riid_t) noexcept;
        #endif

        explicit IComPtr(const IUnknown* iUnknown, REFIID riid = riid_t);

        // COPY CONSTRUCTOR
        IComPtr(const IComPtr& source) noexcept;

        // COPY-LIKE-CONSTRUCTOR (but for differently parameterised templates).
        template <typename ICom_source, REFIID riid_t_source = __uuidof(ICom_source)>
        IComPtr(const IComPtr<ICom_source, riid_t_source>& source);

        // MOVE CONSTRUCTOR
        IComPtr(IComPtr&& source) noexcept;

        // MOVE-LIKE-CONSTRUCTOR (but for differently parameterised templates)
        template <typename ICom_source, REFIID riid_t_source = __uuidof(ICom_source)>
        IComPtr(IComPtr<ICom_source, riid_t_source>&& source);

        // DESTRUCTOR
        ~IComPtr();

        // COPY ASSIGNMENT
        IComPtr& operator=(const IComPtr& source) noexcept;

        // COPY-LIKE-ASSIGNMENT (but for differently parameterised templates).
        template <typename ICom_source, REFIID riid_t_source = __uuidof(ICom_source)>
        IComPtr& operator=(const IComPtr<ICom_source, riid_t_source>& source);

        // MOVE ASSIGNMENT
        IComPtr& operator=(IComPtr&& source) noexcept;

        // MOVE-LIKE-ASSIGNMENT (but for differently parameterised templates).
        template <typename ICom_source, REFIID riid_t_source = __uuidof(ICom_source)>
        IComPtr& operator=(IComPtr<ICom_source, riid_t_source>&& source);

        // ASSIGNMENTs
        #if 0
        // Commented out to enable to parameterize template as follows: IComPtr<IUnknown, IID_IUnknown>.
        IComPtr& operator=(const ICom* source) noexcept;
        #endif

        IComPtr& operator=(const IUnknown* iUnknown) noexcept;

        /// Typecast operator to ICom*.
        explicit operator ICom*() const noexcept;

        /// May be needed, if necessary to fill in the internal ICom pointer from outside, e.g. such as in a call to QueryInterface().
        ICom** operator &();

        #if 0
        ICom& operator *() ... //It is better not to expose this functionality. What should be done with dereferenced pointer to a COM Interface ??? It has no sense.
        #endif

        /// \note Here we prohibit possibility to call AddRef()&Release() with this pointer (by Kenny Kerr: https://msdn.microsoft.com/en-us/magazine/dn904668.aspx)
        IComNoAddRefRelease<ICom>* operator ->();
        bool operator!() const noexcept;

        /// Typecast operator to bool.
        explicit operator bool() const noexcept;

        // METAPROPERTIES
        REFIID GetIID() const noexcept;
        bool& Throws() noexcept;
        HRESULT GetLastHRESULT() const noexcept;

        /// \note This class may misbehave, if, using the returned pointer, COM interface reference counter is manipulated outside IComPtr (damage of invariant).
        ICom* Get() const noexcept;

        // MANIPULATION METHODS
        IComPtr& Reset(const ICom* other = nullptr) noexcept;
        IComPtr& Reset(const IComPtr& other) noexcept;
        IComPtr& Swap(IComPtr& other) noexcept;
        IComPtr& Copy(const ICom* source) noexcept;
        IComPtr& Copy(const IComPtr& source) noexcept;

        /// \note This function does not increment interface reference counter. Use with care.
        IComPtr& Attach(const ICom* other) noexcept;

        /// \note This function does not increment interface reference counter. Use with care.
        IComPtr& Attach(const IComPtr& other) noexcept;

        /// \note This function does not decrement interface reference counter. Use with care.
        ICom* Detach() noexcept;

        /// May be needed, if necessary to fill in the internal ICom pointer from outside, e.g. such as in a call to QueryInterface().
        ICom** GetIComPointerAddress();

        /// Helps to avoid call to QueryInterface(), if duplicated pointer is to the same COM interface, as this one.
        IComPtr& DuplicateIComPointer(ICom** pDuplicatedIComPointer) noexcept;

        /// We can obtain IComPtr to another COM interface.
        /// We can use this IComPtr further, or just once to see, whether the underlying COM component supports the requested COM interface.
        template <typename ICom_another, REFIID riid_t_another = __uuidof(ICom_another)>
        IComPtr<ICom_another, riid_t_another> GetAnotherIComPtr() const;

        /// Just a wrapper: it saves passing two last parameters to the ::CoCreateInstance() function.
        _Check_return_ HRESULT CoCreateInstance(_In_     REFCLSID   rclsid,
                                                _In_opt_ LPUNKNOWN  pUnkOuter,
                                                _In_     DWORD      dwClsContext);

    private:
        // INTERNAL SERVICE METHODS
        ULONG AddRef_internal() noexcept;
        ULONG ReplaceWithNewIComPointer(const ICom* newComInterface = nullptr, bool doAddRef = true ) noexcept;
        ULONG ReplaceWithNewIComPointer(const IComPtr& newComInterface) noexcept;
    };


    // Non-member IComPtr operations&operators.

    template <typename ICom>
    void swap(IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept;
    template <typename ICom>
    bool operator == (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept;
    template <typename ICom>
    bool operator != (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept;

    #if 0
    There is probably no sense in the following comparisons, because the pointers compared are pointers to COM interfaces.
    bool operator < (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept  ...
    bool operator > (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept  ...
    bool operator <= (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept ...
    bool operator >= (IComPtr<ICom>& left, IComPtr<ICom>& right) noexcept ...
    #endif
}

#include "IComPtr.cpp"