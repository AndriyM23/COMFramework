#include "Utils.h"
#include "IComPtr.h"

#define DECLARE_AND_DEFINE_GUIDS
#include "Interfaces.h"

#include <conio.h>

using namespace std;
using namespace Utils;
using namespace COMLibrary;

CommandLineOfProcess cmdLn;
EnvVarsOfProcess     envVars;

int _tmain()
try
{
    ::CoInitialize(NULL);
    {
        HRESULT result;
        IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic01)>   iInterfB01;
        IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic02)>   iInterfB02;
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponent01, , , IInterfBasic01, )"));
        //result = iInterfB01.CoCreateInstance(CLSID_COMComponent01,
        //                                     nullptr,
        //                                     CLSCTX_INPROC_SERVER);
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponent01, ..., IInterfBasic01) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic01"));
        //iInterfB01->F01_IInterfBasic01();
        //iInterfB01->F02_IInterfBasic01();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic04)>   iInterfB04;
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponent04, , , IInterfBasic04, )"));
        //result = iInterfB04.CoCreateInstance(CLSID_COMComponent04,
        //                                     nullptr,
        //                                     CLSCTX_INPROC_SERVER);
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponent04, ..., IInterfBasic04) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic04"));
        //iInterfB04->F01_IInterfBasic04();
        //iInterfB04->F02_IInterfBasic04();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Retrieve set IInterfBasic05 from CLSID_COMComponent01"));
        //IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic05)>   iInterfB05;
        //iInterfB05.Throws() = false;
        //iInterfB05 = iInterfB01;

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB01"));
        //iInterfB01.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB04"));
        //iInterfB04.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponentI040506, , , IInterfBasic04, )"));
        //result = ::CoCreateInstance(CLSID_COMComponentI040506,
        //                            nullptr,
        //                            CLSCTX_INPROC_SERVER,
        //                            iInterfB04.GetIID(),
        //                            reinterpret_cast<LPVOID*>(&iInterfB04));
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponentI040506, ..., IInterfBasic04) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic04"));
        //iInterfB04->F01_IInterfBasic04();
        //iInterfB04->F02_IInterfBasic04();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Retrieve set IInterfBasic05 from CLSID_COMComponentI040506"));
        //iInterfB05 = iInterfB04;

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic05"));
        //iInterfB05->F01_IInterfBasic05();
        //iInterfB05->F02_IInterfBasic05();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB04"));
        //iInterfB04.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB05"));
        //iInterfB05.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponentI010203A04A05A06, , , IInterfBasic06, )"));
        //IComPtr<ICOMPTR_TEMPLATE_ARGUMENTS(IInterfBasic06)>   iInterfB06;
        //result = iInterfB06.CoCreateInstance(CLSID_COMComponentI010203A04A05A06,
        //                                     nullptr,
        //                                     CLSCTX_INPROC_SERVER);
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponentI010203A04A05A06, ..., IInterfBasic06) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic06"));
        //iInterfB06->F01_IInterfBasic06();
        //iInterfB06->F02_IInterfBasic06();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Retrieve IInterfBasic04 from IInterfBasic06(CLSID_COMComponentI010203A04A05A06)"));
        //iInterfB04 = iInterfB06;

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB06"));
        //iInterfB06.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic04"));
        //iInterfB04->F01_IInterfBasic04();
        //iInterfB04->F02_IInterfBasic04();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Retrieve IInterfBasic01 from IInterfBasic04(CLSID_COMComponentI010203A04A05A06)"));
        //iInterfB01 = iInterfB04;

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB04"));
        //iInterfB04.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic01"));
        //iInterfB01->F01_IInterfBasic01();
        //iInterfB01->F02_IInterfBasic01();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB01"));
        //iInterfB01.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();


        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponent02A01, , , IInterfBasic01, )"));
        //result = iInterfB01.CoCreateInstance(CLSID_COMComponent02A01,
        //                                     nullptr,
        //                                     CLSCTX_INPROC_SERVER);
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponent02A01, ..., IInterfBasic01) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic01"));
        //iInterfB01->F01_IInterfBasic01();
        //iInterfB01->F02_IInterfBasic01();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB01"));
        //iInterfB01.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponent02A01, , , IInterfBasic02, )"));
        //result = iInterfB02.CoCreateInstance(CLSID_COMComponent02A01,
        //                                     nullptr,
        //                                     CLSCTX_INPROC_SERVER);
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponent02A01, ..., IInterfBasic02) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic01"));
        //iInterfB02->F01_IInterfBasic02();
        //iInterfB02->F02_IInterfBasic02();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB01"));
        //iInterfB02.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        trace(_TEXT("-------------------------------------------------------------------------------"));
        trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponentI04A02AA01, , , IInterfBasic01, )"));
        result = iInterfB01.CoCreateInstance(CLSID_COMComponentI04A02AA01,
                                             nullptr,
                                             CLSCTX_INPROC_SERVER);
        if (FAILED(result))
            throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponentI04A02AA01, ..., IInterfBasic01) failed."));

        trace(_TEXT("-------------------------------------------------------------------------------"));
        trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic01"));
        iInterfB01->F01_IInterfBasic01();
        iInterfB01->F02_IInterfBasic01();

        trace(_TEXT("-------------------------------------------------------------------------------"));
        trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB01"));
        iInterfB01.Reset();

        trace(_TEXT("-------------------------------------------------------------------------------"));
        trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        ::CoFreeUnusedLibraries();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoCreateInstance(CLSID_COMComponentI04A02AA01, , , IInterfBasic02, )"));
        //result = iInterfB02.CoCreateInstance(CLSID_COMComponentI04A02AA01,
        //    nullptr,
        //    CLSCTX_INPROC_SERVER);
        //if (FAILED(result))
        //    throw WinAPIException(result, _TEXT(__FUNCTION__) _TEXT("|CoCreateInstance(CLSID_COMComponentI04A02AA01, ..., IInterfBasic02) failed."));

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Use interface IInterfBasic02"));
        //iInterfB02->F01_IInterfBasic02();
        //iInterfB02->F02_IInterfBasic02();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Reset iInterfB02"));
        //iInterfB02.Reset();

        //trace(_TEXT("-------------------------------------------------------------------------------"));
        //trace(_TEXT(__FUNCTION__) _TEXT("| Call CoFreeUnusedLibraries()"));
        //::CoFreeUnusedLibraries();

        trace(_TEXT("-------------------------------------------------------------------------------"));
    }

    trace(_TEXT("-------------------------------------------------------------------------------"));
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Calling ::CoUninitialize()"));
    ::CoUninitialize();    // Also does work of ::CoFreeUnusedLibraries();

    trace(_TEXT("--------------------END--------------------------------------------------------"));
    _getch();
    return ERROR_SUCCESS;
}
catch (const WinAPIException& e)
{
    _tstring message;
    e.what(message);
    trace(_TEXT("EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE\n%s"), message.c_str());
    trace(_TEXT(__FUNCTION__) _TEXT("|>>>|Calling ::CoUninitialize()"));
    ::CoUninitialize();

    trace(_TEXT("-------------------------------------------------------------------------------"));
    _getch(); 
    return e.LastError();
}
