#pragma once

#include "WinAPIException.h"

namespace Utils
{
    void trace(const std::_tstring msgstr, ...);


    template<typename FinalFunc>
    struct FinalAction 
    {
        FinalFunc finalFunc;

        FinalAction(FinalFunc func)
            : finalFunc{ func } 
        {
        }

        ~FinalAction()
        { 
            finalFunc(); 
        }
    };

    // Use it like: auto act1 = finally( < [&] { ... } - or other lambda expr. > ); in your code to cleaunp local variables.
    template<class FinalFunc>
    FinalAction<FinalFunc> finally(FinalFunc func)
    {
        return FinalAction<FinalFunc>(func);
    }


    /// Parse version number out of ProgID string: <Program>.<Component>.<Version> .
    /// E.g.: for ProgID string  Visio.Application.3 , the returned result is 3.
    INT GetProgIDVersionNumber(const std::_tstring& progIdStr);


    DWORD SetKeyDefltValToStr(const std::_tstring&  keyName, const std::_tstring&  keyValue);
    std::_tstring GetKeyDefltVal(const std::_tstring&  keyName);
    VOID DeleteKey(HKEY parentKey, const std::_tstring&  keyNameToDelete);


    std::_tstring GuidToString(REFGUID guid);
    GUID StringToGuid(const std::_tstring& guidString);
}

bool operator< (REFGUID guid1, REFGUID guid2);
