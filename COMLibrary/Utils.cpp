#include "Utils.h"

namespace Utils
{
    void trace(const std::_tstring msgstr, ...)
    {
        va_list argptr;
        va_start(argptr, msgstr);
        TCHAR buff[512];
        _vstprintf(buff, msgstr.c_str(), argptr);
        va_end(argptr);
        size_t buffLen = _tcslen(buff);
        buff[buffLen] = _TEXT('\n');
        buff[buffLen + 1] = _TEXT('\0');

        _tcout << buff;
        ::OutputDebugString(buff);
    }

    //================================================================================================
    /// Parse version number out of ProgID string: <Program>.<Component>.<Version> .
    /// E.g.: for ProgID string  Visio.Application.3 , the returned result is 3.
    INT GetProgIDVersionNumber(const std::_tstring& progIdStr)
    {
        std::_tstring::size_type firstPointPos = progIdStr.find_first_of(std::_tstring(_TEXT(".")));
        std::_tstring::size_type lastPointPos = progIdStr.find_last_of(std::_tstring(_TEXT(".")));

        if (lastPointPos == std::_tstring::npos        // There is no point in the ProgID string.
            || lastPointPos == firstPointPos)  // There is only one point in ProgID string
            return (-1);

        std::_tstring versionSubStr = progIdStr.substr(lastPointPos + 1);
        INT verNum = stoi(versionSubStr);
        return verNum;
    }

    //================================================================================================
    DWORD SetKeyDefltValToStr(const std::_tstring&  keyName, const std::_tstring&  keyValue)
    {
        // Open/create key for component's CLSID.
        HKEY    hKey;
        DWORD   disposition;
        LONG    result = RegCreateKeyEx(HKEY_CLASSES_ROOT,
                                        keyName.c_str(),
                                        0,
                                        NULL,
                                        REG_OPTION_NON_VOLATILE,
                                        KEY_ALL_ACCESS,
                                        NULL,
                                        &hKey,
                                        &disposition);
        if (result != ERROR_SUCCESS)
            throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegCreateKeyEx() failed. Key: ")) + keyName);

        if (disposition != REG_OPENED_EXISTING_KEY)
        {
            // The key did not exist before.
            if ((result = RegSetValueEx(hKey,
                                        NULL,
                                        0,
                                        REG_SZ,
                                        reinterpret_cast<const BYTE*>(keyValue.c_str()),
                                        (static_cast<DWORD>(keyValue.length()) + 1) * sizeof(std::_tstring::value_type))) != ERROR_SUCCESS)
            {
                ::RegCloseKey(hKey);
                throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegSetValueEx() failed. Key: ")) + keyName);
            }
        }

        ::RegCloseKey(hKey);
        return disposition;
    }


    std::_tstring GetKeyDefltVal(const std::_tstring&  keyName)
    {
        // Open/create key for component's CLSID.
        HKEY    hKey;
        LONG    result = RegOpenKeyEx(HKEY_CLASSES_ROOT,
            keyName.c_str(),
            0,
            KEY_QUERY_VALUE,
            &hKey);
        if (result != ERROR_SUCCESS)
            throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegOpenKeyEx() failed. Key: ")) + keyName);

        // The key did not exist before.
        BYTE    keyDefaultValue[512];
        DWORD   bytesInBuffer{ ELEMENTS_IN_ARRAY(keyDefaultValue) };
        result = RegQueryValueEx(   hKey,
                                    NULL,
                                    NULL,
                                    NULL,
                                    keyDefaultValue,
                                    &bytesInBuffer);
        ::RegCloseKey(hKey);
        if (result != ERROR_SUCCESS)
            throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegQueryValueEx() failed. Key: ")) + keyName);

        return std::_tstring(reinterpret_cast<PTSTR>(keyDefaultValue));
    }


    VOID DeleteKey(HKEY parentKey, const std::_tstring&  keyNameToDelete)
    {
#if (_WIN32_WINNT >= 0x0600)

        LONG result = RegDeleteTree(parentKey, keyNameToDelete.c_str());
        if (result != ERROR_SUCCESS)
            throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegDeleteTree() failed. Key: ")) + keyNameToDelete);

#else

        HKEY  keyToDelete;
        LONG  result = RegOpenKeyEx(parentKey,
                                    keyNameToDelete.c_str(),
                                    0,
                                    KEY_ALL_ACCESS,
                                    &keyToDelete);
        if (result != ERROR_SUCCESS)
            throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegOpenKeyEx() failed. Key: ")) + keyNameToDelete);

        // Enumerate subkeys and delete them recursively.
        TCHAR regKeyName[512];
        DWORD regKeyNameLen{ ELEMENTS_IN_ARRAY(regKeyName) };
        std::_tstring strRegKeyName;
        while (RegEnumKeyEx(keyToDelete,
                            0,
                            regKeyName,
                            &regKeyNameLen,
                            NULL,
                            NULL,
                            NULL,
                            NULL) == ERROR_SUCCESS  
                 && (strRegKeyName = std::_tstring(regKeyName), strRegKeyName.length() == regKeyNameLen))
        {
            try
            {
                DeleteKey(keyToDelete, strRegKeyName);
            }
            catch (const WinAPIException& e)
            {
                UNREFERENCED_PARAMETER(e);
                ::RegCloseKey(keyToDelete);
                throw;
            }
            regKeyNameLen = ELEMENTS_IN_ARRAY(regKeyName);
        }

        ::RegCloseKey(keyToDelete);

        if ((result = RegDeleteKey(parentKey, keyNameToDelete.c_str())) != ERROR_SUCCESS)
            throw WinAPIException(result, std::_tstring(_TEXT(__FUNCTION__) _TEXT(" RegDeleteKey() failed. Key: ")) + keyNameToDelete);

#endif
    }


    //================================================================================================
    std::_tstring GuidToString(REFGUID guid)
    {
        WCHAR   wguidstr[64]{ L'\0' };
        ::StringFromGUID2(guid, wguidstr, ELEMENTS_IN_ARRAY(wguidstr) - 1);
#ifdef _UNICODE
        return std::_tstring(wguidstr);
#else
        CHAR    guidstr[64]{ '\0' };
        wcstombs(guidstr, wguidstr, ELEMENTS_IN_ARRAY(guidstr) - 1);
        return std::_tstring(guidstr);
#endif
    }

    GUID StringToGuid(const std::_tstring& guidString)
    {
        GUID guid;
#ifdef _UNICODE
        ::CLSIDFromString(guidString.c_str(), &guid);
#else
        WCHAR wguidstr[64]{ L'\0' };
        mbstowcs(wguidstr, guidString.c_str(), ELEMENTS_IN_ARRAY(wguidstr) - 1);
        ::CLSIDFromString(wguidstr, &guid);
#endif
        return guid;
    }
}

bool operator< (REFGUID guid1, REFGUID guid2)
{
    if (guid1.Data1 != guid2.Data1)
        return guid1.Data1 < guid2.Data1;

    if (guid1.Data2 != guid2.Data2)
        return guid1.Data2 < guid2.Data2;

    if (guid1.Data3 != guid2.Data3)
        return guid1.Data3 < guid2.Data3;

    for (int i = 0; i<ELEMENTS_IN_ARRAY(guid2.Data4); i++)
        if (guid1.Data4[i] != guid2.Data4[i])
            return guid1.Data4[i] < guid2.Data4[i];

    return false;
}
