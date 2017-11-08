#include "Utils.h"
#include "COMCategoriesViewer.h"

using namespace std;
using namespace Utils;

namespace COMLibrary
{

void PrintAllComponentCategories(COMCategoriesManager& comcatmgr, COMCategoriesConstRef moreInfoOnTheseCategs)
{
    COMCategories categories = comcatmgr.enumerateAllCategories();
    COMCategoriesCIter iter_cat = categories.cbegin();
    COMCategoriesCIter iter_cat_end = categories.cend();
    set<CLSID> clsids;
    COMCategoriesRef moreInfoOnTheseCategs_noconst = const_cast<COMCategoriesRef>(moreInfoOnTheseCategs);

    trace(_TEXT("Component categories (%d):"), categories.size());
    while (iter_cat != iter_cat_end)
    {
        trace(_TEXT("%s|%s"), iter_cat->str_catid.c_str(), iter_cat->description.c_str());

        COMCategories::iterator foundPos = std::find(moreInfoOnTheseCategs_noconst.begin(), moreInfoOnTheseCategs_noconst.end(), *iter_cat);
        if (foundPos != moreInfoOnTheseCategs_noconst.end())
        {
            *foundPos = *iter_cat;
            COMCategories implementedCategories = { *iter_cat };
            COMCategories requiredCategories = { *iter_cat };
            std::set<CLSID> implementingClsids = comcatmgr.enumClassesOfCategories(implementedCategories, requiredCategories,
                COMCategoriesManager::CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_IMPLEMENTING);
            std::set<CLSID> requiringClsids = comcatmgr.enumClassesOfCategories(implementedCategories, requiredCategories,
                COMCategoriesManager::CLASSES_ENUMERATION_DISPOSITION::ENUMERATE_ONLY_REQUIRING);

            clsids.insert(implementingClsids.begin(), implementingClsids.end());
            clsids.insert(requiringClsids.begin(), requiringClsids.end());

            if (!implementingClsids.empty())
            {
                trace(_TEXT("   Implemented by CLSIDs:"));
                std::set<CLSID>::iterator iter_vec_clsid = implementingClsids.begin();
                while (iter_vec_clsid != implementingClsids.end())
                {
                    COMComponentInfo clsidInfo;
                    clsidInfo.Initialize(*iter_vec_clsid++);
                    trace(_TEXT("   %s|%s"), clsidInfo.str_clsid.c_str(), clsidInfo.friendlyName.c_str());
                }
            }

            if (!requiringClsids.empty())
            {
                trace(_TEXT("   Required by CLSIDs:"));
                std::set<CLSID>::iterator iter_vec_clsid = requiringClsids.begin();
                while (iter_vec_clsid != requiringClsids.end())
                {
                    COMComponentInfo clsidInfo;
                    clsidInfo.Initialize(*iter_vec_clsid++);
                    trace(_TEXT("   %s|%s"), clsidInfo.str_clsid.c_str(), clsidInfo.friendlyName.c_str());
                }
            }
        }
        ++iter_cat;
    }

    if (!clsids.empty())
    {
        trace(_TEXT("\nFollowing CLSIDs:\n"));
        set<CLSID>::const_iterator clsids_iter = clsids.cbegin();
        while (clsids_iter != clsids.cend())
        {
            COMComponentInfo clsidInfo;
            clsidInfo.Initialize(*clsids_iter);
            trace(_TEXT("%s|%s"), clsidInfo.str_clsid.c_str(), clsidInfo.friendlyName.c_str());

            COMCategories implementedCategories;
            COMCategories requiredCategories;

            comcatmgr.enumCategoriesOfClass(*clsids_iter, implementedCategories, requiredCategories);
            if (!implementedCategories.empty())
            {
                trace(_TEXT("   Implements following CATIDs:"));
                COMCategoriesCIter iter_comcat = implementedCategories.cbegin();
                while (iter_comcat != implementedCategories.cend())
                {
                    trace(_TEXT("   %s|%s"), iter_comcat->str_catid.c_str(), iter_comcat->description.c_str());
                    ++iter_comcat;
                }
            }

            if (!requiredCategories.empty())
            {
                trace(_TEXT("   Requires following CATIDs:"));
                COMCategoriesCIter iter_comcat = requiredCategories.cbegin();
                while (iter_comcat != requiredCategories.cend())
                {
                    trace(_TEXT("   %s|%s"), iter_comcat->str_catid.c_str(), iter_comcat->description.c_str());
                    ++iter_comcat;
                }
            }

            ++clsids_iter;
        }
    }
}

}
