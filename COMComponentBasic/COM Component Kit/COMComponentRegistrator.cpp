#include "Utils.h"
#include "COMComponentRegistrator.h"

#include <string.h>     //for _tcscpy().
#include <algorithm>    //for std::set_difference().

using namespace std;
using namespace Utils;

namespace COMLibrary
{

COMCategory::COMCategory(const CATEGORYINFO& info)
{
    memcpy(&this->info, &info, sizeof(CATEGORYINFO));
    str_catid = GuidToString(info.catid);
    #if (defined(_UNICODE) && !defined(OLE2ANSI) || !defined(_UNICODE) && defined(OLE2ANSI))
        description = _tstring(info.szDescription);
    #elif (!defined(_UNICODE) && !defined(OLE2ANSI))
    {
        CHAR componentInformation[128];
        wcstombs(componentInformation, info.szDescription, 128);
        description = string(componentInformation);
    }
    #else
    {
        WCHAR componentInformation[128];
        mbstowcs(componentInformation, info.szDescription, 128);
        description = wstring(componentInformation);
    }
    #endif
}

COMCategory::COMCategory(const std::_tstring& str_catid, const std::_tstring& description)
    : str_catid(str_catid),
      description(description)
{
    info.catid = StringToGuid(str_catid);
    info.lcid = ::GetUserDefaultLCID();
    #if (defined(_UNICODE) && !defined(OLE2ANSI) || !defined(_UNICODE) && defined(OLE2ANSI))
        _tcscpy_s(info.szDescription, 128, description.c_str());
        info.szDescription[127] = _TEXT('\0');
    #elif (!defined(_UNICODE) && !defined(OLE2ANSI))
        mbstowcs(info.szDescription, description.c_str(), 128);
        info.szDescription[127] = L'\0';
    #else
        wcstombs(info.szDescription, description.c_str(), 128);
        info.szDescription[127] = '\0';
    #endif
}

COMCategory::COMCategory(REFCATID catid, const std::_tstring& description)
    : description(description)
{
    info.catid = catid;
    str_catid = GuidToString(catid);
    info.lcid = ::GetUserDefaultLCID();
    #if (defined(_UNICODE) && !defined(OLE2ANSI) || !defined(_UNICODE) && defined(OLE2ANSI))
        _tcscpy_s(info.szDescription, 128, description.c_str());
        info.szDescription[127] = _TEXT('\0');
    #elif (!defined(_UNICODE) && !defined(OLE2ANSI))
        mbstowcs(info.szDescription, description.c_str(), 128);
        info.szDescription[127] = L'\0';
    #else
        wcstombs(info.szDescription, description.c_str(), 128);
        info.szDescription[127] = '\0';
    #endif
}

void COMCategory::setDescription(const std::wstring& wDescription)
{
    #ifndef OLE2ANSI
    wcscpy_s(info.szDescription, 128, wDescription.c_str());
    info.szDescription[127] = L'\0';
    #else
    wcstombs(comCategory.info.szDescription, wDescription.c_str(), 128);
    info.szDescription[127] = '\0';
    #endif

    #ifdef _UNICODE
    description = wDescription;
    #else
    {
        CHAR descriptionArr[128];
        wcstombs(descriptionArr, wDescription.c_str(), 128);
        descriptionArr[127] = '\0';
        description = string(descriptionArr);
    }
    #endif
}

bool COMCategory::operator == (const COMCategory& comCategory)
{
    return static_cast<bool>(::IsEqualCATID(info.catid, comCategory.info.catid));
}

//====================================================================================================

COMCategoriesManager::COMCategoriesManager()
{
    try
    {
        ::CoInitialize(NULL);
        obtainICatInformation();
        obtainICatRegister();
    }
    catch (const WinAPIException& e)
    {
        UNREFERENCED_PARAMETER(e);
        ::CoUninitialize();
        throw;
    }
}

COMCategoriesManager::~COMCategoriesManager()
{
    ::CoUninitialize();
}

COMCategories COMCategoriesManager::enumerateAllCategories()
{
    IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IEnumCATEGORYINFO)> pIEnumCATEGORYINFO;
    HRESULT result = m_pICatInfromation->EnumCategories(::GetUserDefaultLCID(), &pIEnumCATEGORYINFO);
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatInfromation->EnumCategories() failed.")));

    COMCategories categories;
    do
    {
        CATEGORYINFO    catInfo[10];
        ULONG           itemsRetrieved = 0;
        result = pIEnumCATEGORYINFO->Next(10, catInfo, &itemsRetrieved);
        for (UINT i = 0; i < itemsRetrieved; ++i)
        {
            COMCategory cat(catInfo[i]);
            categories.push_back(cat);
        }
    } while (result == S_OK);

    return categories;
}

void COMCategoriesManager::registerCategories(COMCategoriesConstRef newCategories)
{
    unique_ptr<CATEGORYINFO[]> pCatInfo = COMCategoriesToCATEGORYINFOArr(newCategories);
    if (!pCatInfo)
        throw WinAPIException(E_INVALIDARG, _tstring(_TEXT(__FUNCTION__) _TEXT("|newCategories are not specified. Nothing to register.")));
    HRESULT result = m_pICatRegister->RegisterCategories(newCategories.size(), pCatInfo.get());
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatRegister->RegisterCategories() failed.")));
}

void COMCategoriesManager::unregisterCategories(COMCategoriesConstRef existingCategories)
{
    unique_ptr<CATID[]> pCatID = COMCategoriesToCATIDArr(existingCategories);
    if (!pCatID)
        throw WinAPIException(E_INVALIDARG, _tstring(_TEXT(__FUNCTION__) _TEXT("|existingCategories are not specified. Nothing to unregister.")));
    HRESULT result = m_pICatRegister->UnRegisterCategories(existingCategories.size(), pCatID.get());
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatRegister->UnRegisterCategories() failed.")));
}

void COMCategoriesManager::registerClassImplementingCategories(REFCLSID clsid, COMCategoriesConstRef implementedCategories)
{
    unique_ptr<CATID[]> pCatID = COMCategoriesToCATIDArr(implementedCategories);
    if (!pCatID)
        throw WinAPIException(E_INVALIDARG, _tstring(_TEXT(__FUNCTION__) _TEXT("|implementedCategories are not specified. Nothing to register.")));
    HRESULT result = m_pICatRegister->RegisterClassImplCategories(clsid, implementedCategories.size(), pCatID.get());
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatRegister->RegisterClassImplCategories() failed.")));
}

void COMCategoriesManager::registerClassRequiringCategories(REFCLSID clsid, COMCategoriesConstRef requiredCategories)
{
    unique_ptr<CATID[]> pCatID = COMCategoriesToCATIDArr(requiredCategories);
    if (!pCatID)
        throw WinAPIException(E_INVALIDARG, _tstring(_TEXT(__FUNCTION__) _TEXT("|requiredCategories are not specified. Nothing to register.")));
    HRESULT result = m_pICatRegister->RegisterClassReqCategories(clsid, requiredCategories.size(), pCatID.get());
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatRegister->RegisterClassReqCategories() failed.")));
}

void COMCategoriesManager::unregisterClassImplementingCategories(REFCLSID clsid, COMCategoriesConstRef implementedCategories)
{
    unique_ptr<CATID[]> pCatID = COMCategoriesToCATIDArr(implementedCategories);
    if (!pCatID)
        throw WinAPIException(E_INVALIDARG, _tstring(_TEXT(__FUNCTION__) _TEXT("|implementedCategories are not specified. Nothing to unregister.")));
    HRESULT result = m_pICatRegister->UnRegisterClassImplCategories(clsid, implementedCategories.size(), pCatID.get());
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatRegister->UnRegisterClassImplCategories() failed.")));
}

void COMCategoriesManager::unregisterClassRequiringCategories(REFCLSID clsid, COMCategoriesConstRef requiredCategories)
{
    unique_ptr<CATID[]> pCatID = COMCategoriesToCATIDArr(requiredCategories);
    if (!pCatID)
        throw WinAPIException(E_INVALIDARG, _tstring(_TEXT(__FUNCTION__) _TEXT("|requiredCategories are not specified. Nothing to unregister.")));
    HRESULT result = m_pICatRegister->UnRegisterClassReqCategories(clsid, requiredCategories.size(), pCatID.get());
    if (FAILED(result))
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|m_pICatRegister->UnRegisterClassReqCategories() failed.")));
}

set<CLSID> COMCategoriesManager::enumClassesOfCategories (  COMCategoriesConstRef implementedCategories,
                                                            COMCategoriesConstRef requiredCategories,
                                                            CLASSES_ENUMERATION_DISPOSITION enumDisposition )
{
    set<CLSID> sClasses;
    if (implementedCategories.empty())
        return sClasses;

    if (requiredCategories.empty() && enumDisposition == CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_REQUIRING)
        return sClasses;

    unique_ptr<CATID[]> pImplementedCatID = COMCategoriesToCATIDArr(implementedCategories);
    unique_ptr<CATID[]> pRequiredCatID = COMCategoriesToCATIDArr(requiredCategories);

    INT implementedCategoriesSize;
    INT requiredCategoriesSize;
    CATID* pImplementedCategoryIDs;
    CATID* pRequiredCategoryIDs;
    IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IEnumCLSID)> pIEnumCLSID;
    switch (enumDisposition)
    {
    case CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_IMPLEMENTING:
        implementedCategoriesSize = implementedCategories.size();
        requiredCategoriesSize = -1;
        pImplementedCategoryIDs = pImplementedCatID.get();
        pRequiredCategoryIDs = nullptr;
        break;
    case CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_IMPLEMENTING_AND_REQUIRING:
    case CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_REQUIRING:
        implementedCategoriesSize = implementedCategories.size();
        requiredCategoriesSize = requiredCategories.size();
        pImplementedCategoryIDs = pImplementedCatID.get();
        pRequiredCategoryIDs = pRequiredCatID.get();
        break;
    }
    HRESULT result = m_pICatInfromation->EnumClassesOfCategories(implementedCategoriesSize,
                                                                 pImplementedCategoryIDs,
                                                                 requiredCategoriesSize,
                                                                 pRequiredCategoryIDs,
                                                                 &pIEnumCLSID);
    if (SUCCEEDED(result))
    {
        do
        {
            CLSID clsidarr[10];
            ULONG itemsRetrieved;
            result = pIEnumCLSID->Next(10, clsidarr, &itemsRetrieved);
            for (UINT i = 0; i<itemsRetrieved; ++i)
                sClasses.insert(clsidarr[i]);
        } while (result == S_OK);

        if (enumDisposition == CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_REQUIRING)
        {
            set<CLSID> sClassesImplOnly;
            set<CLSID> sClassesReqOnly;

            // Parameters are like for CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_IMPLEMENTING
            implementedCategoriesSize = implementedCategories.size();
            requiredCategoriesSize = -1;
            pImplementedCategoryIDs = pImplementedCatID.get();
            pRequiredCategoryIDs = nullptr;

            pIEnumCLSID.Reset();
            result = m_pICatInfromation->EnumClassesOfCategories(implementedCategoriesSize,
                                                                 pImplementedCategoryIDs,
                                                                 requiredCategoriesSize,
                                                                 pRequiredCategoryIDs,
                                                                 &pIEnumCLSID);
            if (SUCCEEDED(result))
            {
                do
                {
                    CLSID clsidarr[10];
                    ULONG itemsRetrieved;
                    result = pIEnumCLSID->Next(10, clsidarr, &itemsRetrieved);
                    for (UINT i = 0; i<itemsRetrieved; ++i)
                        sClassesImplOnly.insert(clsidarr[i]);
                } while (result == S_OK);

                // Find difference between CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_IMPLEMENTING_AND_REQUIRING and CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_IMPLEMENTING
                vector<CLSID> vClassesReqOnly(sClasses.size());
                vector<CLSID>::iterator lastDifference = set_difference(sClasses.begin(), sClasses.end(),
                                                                        sClassesImplOnly.begin(), sClassesImplOnly.end(),
                                                                        vClassesReqOnly.begin(), std::less<GUID>());
                sClasses.clear();
                sClasses.insert(vClassesReqOnly.begin(), lastDifference);
            }
        }
    }

    return sClasses;
}

void COMCategoriesManager::enumCategoriesOfClass(REFCLSID           clsid,
                                                 COMCategoriesRef   implementedCategories,
                                                 COMCategoriesRef   requiredCategories)
{
    implementedCategories.clear();
    requiredCategories.clear();

    //------------------------------------------------
    IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IEnumCATID)> pIEnumCATID;
    HRESULT result = m_pICatInfromation->EnumImplCategoriesOfClass(clsid, &pIEnumCATID);
    if (SUCCEEDED(result))
    {
        do
        {
            CATID catidarr[10];
            ULONG itemsRetrieved;
            result = pIEnumCATID->Next(10, catidarr, &itemsRetrieved);
            for (UINT i = 0; i<itemsRetrieved; ++i)
            {
                COMCategory comCategory{ catidarr[i] };

                LPWSTR pszCategoryDescription;
                HRESULT resultGetDescr = m_pICatInfromation->GetCategoryDesc(catidarr[i], ::GetUserDefaultLCID(), &pszCategoryDescription);
                if (SUCCEEDED(resultGetDescr))
                {
                    comCategory.setDescription(pszCategoryDescription);
                    ::CoTaskMemFree(pszCategoryDescription);
                }

                implementedCategories.push_back(comCategory);
            }
        } while (result == S_OK);
    }

    //------------------------------------------------
    pIEnumCATID.Reset();
    result = m_pICatInfromation->EnumReqCategoriesOfClass(clsid, &pIEnumCATID);
    if (SUCCEEDED(result))
    {
        do
        {
            CATID catidarr[10];
            ULONG itemsRetrieved;
            result = pIEnumCATID->Next(10, catidarr, &itemsRetrieved);
            for (UINT i = 0; i<itemsRetrieved; ++i)
            {
                COMCategory comCategory{ catidarr[i] };

                LPWSTR pszCategoryDescription;
                HRESULT resultGetDescr = m_pICatInfromation->GetCategoryDesc(catidarr[i], ::GetUserDefaultLCID(), &pszCategoryDescription);
                if (SUCCEEDED(resultGetDescr))
                {
                    comCategory.setDescription(pszCategoryDescription);
                    ::CoTaskMemFree(pszCategoryDescription);
                }

                requiredCategories.push_back(comCategory);
            }
        } while (result == S_OK);
    }
}

void COMCategoriesManager::obtainICatInformation()
{
    HRESULT result = ::CoCreateInstance(CLSID_StdComponentCategoriesMgr,
                                        NULL,
                                        CLSCTX_ALL,
                                        m_pICatInfromation.GetIID(),
                                        reinterpret_cast<void**>(&m_pICatInfromation));
    if (FAILED(result))
    {
        m_pICatInfromation.Reset();
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_StdComponentCategoriesMgr,,,IID_ICatInformation,) failed.")));
    }
}

void COMCategoriesManager::obtainICatRegister()
{
    HRESULT result = ::CoCreateInstance(CLSID_StdComponentCategoriesMgr,
                                        NULL,
                                        CLSCTX_ALL,
                                        m_pICatRegister.GetIID(),
                                        reinterpret_cast<void**>(&m_pICatRegister));
    if (FAILED(result))
    {
        m_pICatRegister.Reset();
        throw WinAPIException(result, _tstring(_TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_StdComponentCategoriesMgr,,,IID_ICatRegister,) failed.")));
    }
}

std::unique_ptr<CATID[]> COMCategoriesManager::COMCategoriesToCATIDArr(COMCategoriesConstRef comCategories)
{
    unique_ptr<CATID[]> pCatID;
    if (comCategories.empty())
        return pCatID;
    pCatID = std::make_unique<CATID[]>(comCategories.size());
    COMCategoriesCIter iter = comCategories.cbegin();
    COMCategoriesCIter end_iter = comCategories.cend();
    UINT i = 0;
    while (iter != end_iter)
        pCatID[i++] = iter++->info.catid;

    return pCatID;
}

std::unique_ptr<CATEGORYINFO[]> COMCategoriesManager::COMCategoriesToCATEGORYINFOArr(COMCategoriesConstRef comCategories)
{
    unique_ptr<CATEGORYINFO[]> pCatInfo;
    if (comCategories.empty())
        return pCatInfo;    
    pCatInfo = std::make_unique<CATEGORYINFO[]>(comCategories.size());
    COMCategoriesCIter iter = comCategories.cbegin();
    COMCategoriesCIter end_iter = comCategories.cend();
    UINT i = 0;
    while (iter != end_iter)
        memcpy(pCatInfo.get() + i++, &iter++->info, sizeof(CATEGORYINFO));
    return pCatInfo;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////

static COMCategoriesManager  comCatMgr;

COMComponentInfo::COMComponentInfo()
{
}

COMComponentInfo::COMComponentInfo(REFCLSID clsid)
{
    Initialize(clsid);
}

COMComponentInfo::COMComponentInfo(HMODULE                hModule,
                                   REFCLSID               clsid,
                                   const std::_tstring    &friendlyName,
                                   const std::_tstring    &progId,
                                   const std::_tstring    &versionIndependentProgId,
                                   COMCategoriesConstRef  implementedCategories,
                                   COMCategoriesConstRef  requiredCategories)
    : hModule(hModule),
      clsid(clsid),
      friendlyName(friendlyName),
      progId(progId),
      versionIndependentProgId(versionIndependentProgId),
      implementedCategories(implementedCategories),
      requiredCategories(requiredCategories)
{
    str_clsid = GuidToString(clsid);

    TCHAR componentDllFilenameArr[512];
    ::GetModuleFileName(hModule, componentDllFilenameArr, 512);
    componentDllFilename = _tstring(componentDllFilenameArr);

    currVersionNumber = GetProgIDVersionNumber(progId);
}

void COMComponentInfo::Initialize(REFCLSID clsid)
{
    this->clsid = clsid;
    str_clsid = GuidToString(clsid);

    // CLSID
    _tstring clsidKeyName(_TEXT("CLSID\\"));
    clsidKeyName += str_clsid;
    friendlyName = GetKeyDefltVal(clsidKeyName);

    _tstring inprocServer32SubkeyName = clsidKeyName + _TEXT("\\InprocServer32");
    componentDllFilename = GetKeyDefltVal(inprocServer32SubkeyName);

    _tstring progIdSubkeyName = clsidKeyName + _TEXT("\\ProgID");
    progId = GetKeyDefltVal(progIdSubkeyName);
    currVersionNumber = GetProgIDVersionNumber(progId);

    _tstring verIndepProgIdSubkeyName = clsidKeyName + _TEXT("\\VersionIndependentProgID");
    versionIndependentProgId = GetKeyDefltVal(verIndepProgIdSubkeyName);

    comCatMgr.enumCategoriesOfClass(clsid, implementedCategories, requiredCategories);
}

void COMComponentInfo::Register()
try
{
    // CLSID
    _tstring clsidKeyName(_TEXT("CLSID\\"));
    clsidKeyName += str_clsid;

    DWORD disposition = SetKeyDefltValToStr(clsidKeyName, friendlyName);
    if (disposition == REG_OPENED_EXISTING_KEY)
    {
        // We already have this component registered.
        // TODO: possibly we are trying to install another version of COM component.
        // TODO: identify, which actions have to be done.
        // Unregister();
    }

    _tstring inprocServer32SubkeyName = clsidKeyName + _TEXT("\\InprocServer32");
    SetKeyDefltValToStr(inprocServer32SubkeyName, componentDllFilename);

    _tstring progIdSubkeyName = clsidKeyName + _TEXT("\\ProgID");
    SetKeyDefltValToStr(progIdSubkeyName, progId);

    _tstring verIndepProgIdSubkeyName = clsidKeyName + _TEXT("\\VersionIndependentProgID");
    SetKeyDefltValToStr(verIndepProgIdSubkeyName, versionIndependentProgId);

    // ProgID
    _tstring progIdKeyName = progId;
    SetKeyDefltValToStr(progIdKeyName, friendlyName);

    _tstring progIdClsidKeyName = progIdKeyName + _TEXT("\\CLSID");
    SetKeyDefltValToStr(progIdClsidKeyName, str_clsid);

    // VersionIndependentProgID
    _tstring verIndepProgIdKeyName = versionIndependentProgId;
    SetKeyDefltValToStr(verIndepProgIdKeyName, friendlyName);

    _tstring verIndepProgIdClsidSubkeyName = verIndepProgIdKeyName + _TEXT("\\CLSID");
    SetKeyDefltValToStr(verIndepProgIdClsidSubkeyName, str_clsid);

    _tstring verIndepProgIdCurVerSubkeyName = verIndepProgIdKeyName + _TEXT("\\CurVer");
    SetKeyDefltValToStr(verIndepProgIdCurVerSubkeyName, progId);

    // Register the component as the one, which implements following categories:
    comCatMgr.registerClassImplementingCategories(clsid, implementedCategories);

    // Register the component as the one, which requires following categories:
    comCatMgr.registerClassRequiringCategories(clsid, requiredCategories);
}
catch (const WinAPIException& e)
{
    if (e.LastError() != E_INVALIDARG)
        throw;
}

void COMComponentInfo::Unregister()
try
{
    // UnRegister the component as the one, which implements following categories:
    comCatMgr.unregisterClassImplementingCategories(clsid, implementedCategories);

    // UnRegister the component as the one, which requires following categories:
    comCatMgr.unregisterClassRequiringCategories(clsid, requiredCategories);

    _tstring clsidKeyName(_TEXT("CLSID\\"));
    clsidKeyName += str_clsid;
    DeleteKey(HKEY_CLASSES_ROOT, clsidKeyName);

    _tstring progIdKeyName = progId;
    DeleteKey(HKEY_CLASSES_ROOT, progIdKeyName);

    _tstring verIndepProgIdKeyName = versionIndependentProgId;
    DeleteKey(HKEY_CLASSES_ROOT, verIndepProgIdKeyName);
}
catch (const WinAPIException& e)
{
    _tstring clsidKeyName(_TEXT("CLSID\\"));
    clsidKeyName += str_clsid;
    try
    {
        DeleteKey(HKEY_CLASSES_ROOT, clsidKeyName);
    }
    catch (const WinAPIException&)
    { }

    _tstring progIdKeyName = progId;
    try
    {
        DeleteKey(HKEY_CLASSES_ROOT, progIdKeyName);
    }
    catch (const WinAPIException&)
    { }

    _tstring verIndepProgIdKeyName = versionIndependentProgId;
    try
    {
        DeleteKey(HKEY_CLASSES_ROOT, verIndepProgIdKeyName);
    }
    catch (const WinAPIException&)
    { }

    if (e.LastError() != E_INVALIDARG)
        throw;
}

}
