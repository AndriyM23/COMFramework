#pragma once

#include "IComPtr.h"
#include <ComCat.h>     // Header of Component Categories Manager
#include <set>

namespace COMLibrary
{

struct COMCategory
{
    CATEGORYINFO info;
    std::_tstring  str_catid;
    std::_tstring  description;
    COMCategory(const CATEGORYINFO& info);
    COMCategory(const std::_tstring& str_catid, const std::_tstring& description = _TEXT(""));
    COMCategory(REFCATID catid, const std::_tstring& description = _TEXT(""));
    void setDescription(const std::wstring& wDescription);
    bool operator == (const COMCategory& comCategory);
};

typedef std::vector<COMCategory> COMCategories;
typedef std::vector<COMCategory>& COMCategoriesRef;
typedef const std::vector<COMCategory> COMCategoriesConst;
typedef const std::vector<COMCategory>& COMCategoriesConstRef;
typedef std::vector<COMCategory>::iterator COMCategoriesIter;
typedef std::vector<COMCategory>::const_iterator COMCategoriesCIter;

class COMCategoriesManager
{
    IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(ICatInformation)>    m_pICatInfromation;
    IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(ICatRegister)>       m_pICatRegister;

public:
    COMCategoriesManager();
    ~COMCategoriesManager();

    COMCategories enumerateAllCategories();

    void registerCategories(COMCategoriesConstRef newCategories);
    void unregisterCategories(COMCategoriesConstRef existingCategories);

    void registerClassImplementingCategories(REFCLSID clsid, COMCategoriesConstRef implementedCategories);
    void registerClassRequiringCategories(REFCLSID clsid, COMCategoriesConstRef requiredCategories);

    void unregisterClassImplementingCategories(REFCLSID clsid, COMCategoriesConstRef implementedCategories);
    void unregisterClassRequiringCategories(REFCLSID clsid, COMCategoriesConstRef requiredCategories);

    enum class CLASSES_ENUMERATION_DISPOSITION
    {
        ENUMERATE_ONLY_IMPLEMENTING,
        ENUMERATE_ONLY_REQUIRING,
        ENUMERATE_IMPLEMENTING_AND_REQUIRING  // Conjunction of sets: ENUMERATE_ONLY_IMPLEMENTING & ENUMERATE_ONLY_REQUIRING
    };
    std::set<CLSID> enumClassesOfCategories(COMCategoriesConstRef implementedCategories,
                                            COMCategoriesConstRef requiredCategories,
                                            CLASSES_ENUMERATION_DISPOSITION enumDisposition = CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_IMPLEMENTING_AND_REQUIRING);

    void enumCategoriesOfClass(REFCLSID             clsid,
                               COMCategoriesRef     implementedCategories,
                               COMCategoriesRef     requiredCategories);

private:
    void obtainICatInformation();
    void obtainICatRegister();

    std::unique_ptr<CATID[]> COMCategoriesToCATIDArr(COMCategoriesConstRef implementedCategories);
    std::unique_ptr<CATEGORYINFO[]> COMCategoriesToCATEGORYINFOArr(COMCategoriesConstRef implementedCategories);
};

////////////////////////////////////////////////////////////////////////////////////////////////////////

struct COMComponentInfo
{
    CLSID           clsid;
    std::_tstring   str_clsid;
    std::_tstring   friendlyName;
    std::_tstring   progId;
    std::_tstring   versionIndependentProgId;
    std::_tstring   componentDllFilename;
    INT             currVersionNumber {-1};
    HMODULE         hModule {nullptr};

    COMCategories   implementedCategories;
    COMCategories   requiredCategories;

    COMComponentInfo();
    COMComponentInfo(REFCLSID               clsid);
    COMComponentInfo(HMODULE                hModule,
                     REFCLSID               clsid,
                     const std::_tstring    &friendlyName,
                     const std::_tstring    &progId,
                     const std::_tstring    &versionIndependentProgId,
                     COMCategoriesConstRef  implementedCategories,
                     COMCategoriesConstRef  requiredCategories);

    void Initialize(REFCLSID clsid);

    /// \note  Surround Register() call with try-catch block.
    void Register();
    /// \note  Surround Unregister() call with try-catch block, especially in other try-catch blocks.
    /// \note  Unregister steps are not done in the reverse-order, as informatic theory usually presciribes.
    ///        Unregister steps are done in the same order, as register steps. This smiplifies using these functions together with exceptions.
    ///        If Register() fails and throws an exception, in exception handler it is sufficient to call Unregister() for removing side-effects, 
    ///        left by incomplete Register(). So, please place Unregister() calls for components in the same order, as you Register() them.
    void Unregister();
};
typedef COMComponentInfo *PCOMComponentInfo;
typedef std::shared_ptr<COMComponentInfo> COMComponentInfoPtr;
typedef std::vector<COMComponentInfoPtr> COMComponentsInfo;

}