#pragma once

#include "WindowsUtils.h"
#include <exception>        // For C++ exception handling
#include <stdexcept>
#include <memory>           // For C++ smart pointers.

namespace Utils
{

class WinAPIException : public std::exception
{
    DWORD           lastError;
    std::_tstring   wholeMessage;
    #ifdef _UNICODE
    std::unique_ptr<CHAR[]> wholeMessageArrPtr;
    #endif

    static const DWORD  NONEXISTING_LAST_ERROR { 0xFFFFFFFFUL };

public:
    explicit WinAPIException(const std::_tstring userMessage, ...) noexcept
    {
        lastError = ::GetLastError();
        va_list variadicArgList;
        va_start(variadicArgList, userMessage);
        Initialize(userMessage, variadicArgList);
    }

    explicit WinAPIException(DWORD lastError, const std::_tstring userMessage, ...) noexcept
        : lastError(lastError)
    {
        va_list variadicArgList;
        va_start(variadicArgList, userMessage);
        Initialize(userMessage, variadicArgList);
    }

    WinAPIException(const WinAPIException& source) noexcept
        :   lastError(source.lastError),
            wholeMessage(source.wholeMessage)
    {
        #ifdef _UNICODE
        wholeMessageArrPtr.reset(new CHAR[wholeMessage.length() + 1]{'\0'});
        if (wholeMessageArrPtr)
            strcpy(wholeMessageArrPtr.get(), source.wholeMessageArrPtr.get());
        #endif
    }

    WinAPIException(WinAPIException&& source) noexcept
        :   lastError(source.lastError)
    {
        source.lastError = NONEXISTING_LAST_ERROR;      // Leaving the source in a "moved-from" state.
        wholeMessage = std::move(source.wholeMessage);  // Leaving the source in a "moved-from" state.
        #ifdef _UNICODE
        wholeMessageArrPtr = std::move(source.wholeMessageArrPtr); // Leaving the source in a "moved-from" state.
        #endif
    }

    WinAPIException& operator=(const WinAPIException& source) noexcept
    {
        if (&source != this)
        {
            lastError = source.lastError;
            wholeMessage = source.wholeMessage;
            #ifdef _UNICODE
            wholeMessageArrPtr.reset(new CHAR[wholeMessage.length() + 1]{'\0'});
            if (wholeMessageArrPtr)
                strcpy(wholeMessageArrPtr.get(), source.wholeMessageArrPtr.get());
            #endif
        }
        return *this;
    }

    WinAPIException& operator=(WinAPIException&& source) noexcept
    {
        if (&source != this)
        {
            lastError = source.lastError;
            source.lastError = NONEXISTING_LAST_ERROR;      // Leaving the source in a "moved-from" state.
            wholeMessage = std::move(source.wholeMessage);  // Leaving the source in a "moved-from" state.
            #ifdef _UNICODE
            wholeMessageArrPtr = std::move(source.wholeMessageArrPtr); // Leaving the source in a "moved-from" state.
            #endif
        }
        return *this;
    }

    virtual const char* what() const noexcept
    {
        #ifdef _UNICODE
        return wholeMessageArrPtr.get();
        #else
        return wholeMessage.c_str();
        #endif
    }

    virtual void what(_Out_ std::_tstring& message) const noexcept
    {
        message = wholeMessage;
    }

    DWORD LastError() const noexcept
    {
        return lastError;
    }

private:
    static const DWORD  MAX_MESSAGE_SIZE{ 1024 };

    /// \note This function closes variadicArgList. So, in case of further usage, make a copy first.
    void Initialize(const std::_tstring& userMessage, va_list variadicArgList) noexcept
    {
        TCHAR localCharArr[MAX_MESSAGE_SIZE];
        DWORD fmtMsgResult = ::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM |
                                             FORMAT_MESSAGE_IGNORE_INSERTS |
                                             FORMAT_MESSAGE_MAX_WIDTH_MASK,
                                             NULL,
                                             lastError,
                                             MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                             reinterpret_cast<LPTSTR>(localCharArr),
                                             MAX_MESSAGE_SIZE,
                                             nullptr );
        std::_tstring errorCodeDescription;
        if (fmtMsgResult > 0)
            errorCodeDescription = std::_tstring{localCharArr};

        _vstprintf(localCharArr, userMessage.c_str(), variadicArgList);
        va_end(variadicArgList);
        wholeMessage = std::_tstring(_TEXT("!!!WinAPIException!!! ")) + std::_tstring(localCharArr);
        _stprintf(localCharArr, _TEXT(" LastError=%d=%08.8Xh="), lastError, lastError);
        wholeMessage += std::_tstring(localCharArr) + errorCodeDescription;

        #ifdef _UNICODE
        wholeMessageArrPtr.reset(new CHAR[wholeMessage.length()+1]{'\0'});
        if (wholeMessageArrPtr)
            wcstombs(wholeMessageArrPtr.get(), wholeMessage.c_str(), wholeMessage.length());
        #endif
    }
};

}