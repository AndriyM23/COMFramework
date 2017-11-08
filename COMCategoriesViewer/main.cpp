#include "Utils.h"
#include "COMCategoriesViewer.h"

#define DECLARE_AND_DEFINE_GUIDS
#include "Interfaces.h"

#include <conio.h> // for _getch();

using namespace std;
using namespace COMLibrary;
using namespace Utils;


int _tmain(int argc, PTSTR argv[])
try
{
    bool bUnregisterCategories;
    bool bViewOnly = true;
    if (argc > 1)
    {
        _tstring option_str(argv[1]);
        if (option_str == _TEXT("/u"))
        {
            bUnregisterCategories = true;
            bViewOnly = false;
        }
        if (option_str == _TEXT("/r"))
        {
            bUnregisterCategories = false;
            bViewOnly = false;
        }
    }

    //Create array with component categories to register/unregister/display extended info.
    COMCategories categories{ { COMCategoryIFB01, _TEXT(COMCATEGORYIFB01_DESCRIPTION) },
                              { COMCategoryIFB04, _TEXT(COMCATEGORYIFB04_DESCRIPTION) },
                              { COMCategoryIFB040506, _TEXT(COMCATEGORYIFB040506_DESCRIPTION) },
                              { COMCategoryIFB010203040506, _TEXT(COMCATEGORYIFB010203040506_DESCRIPTION)  } };

    //COMCategories categories{ { COMCategoryIFB01, _TEXT(COMCATEGORYIFB01_DESCRIPTION) } };


    COMCategoriesManager comcatmgr;
    if (bViewOnly)
    {
        PrintAllComponentCategories(comcatmgr, categories);
        _getch();
        return ERROR_SUCCESS;
    }
    else
        PrintAllComponentCategories(comcatmgr);

    if (bUnregisterCategories)
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|--------------------- Unregister component categories --------------------------"));
        comcatmgr.unregisterCategories(categories);
        PrintAllComponentCategories(comcatmgr);
    }
    else
    {
        trace(_TEXT(__FUNCTION__) _TEXT("|--------------------- Register component categories --------------------------"));
        comcatmgr.registerCategories(categories);
        PrintAllComponentCategories(comcatmgr, categories);
    }

    _getch();
    return ERROR_SUCCESS;
}
catch (const WinAPIException& e)
{
    _tstring message;
    e.what(message);
    trace(_TEXT("\nEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE\n%s"), message.c_str());
    _getch();
    return e.LastError();
}
