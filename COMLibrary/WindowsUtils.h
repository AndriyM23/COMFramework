#pragma     once    //Specifies that the file will be included (opened) only once by the compiler when compiling a source code file.

//Before the inclusion of standard Windows headers, we must specify, what platform are we compiling for.
//These macros enables us to use the compiler to detect whether our application uses functions that are not supported 
//on application's target version(s) of Windows.
#undef    _WIN32_WINNT      // See sdkddkver.h
#define   _WIN32_WINNT                              0x0502    //The application works on WXP SP2, Windows Sever 2003 or higher platforms.
//#define   _WIN32_WINNT                              0x0600    //The application works on Windows Vista or higher platforms.
#undef    WINVER            // See sdkddkver.h
#define   WINVER                                    0x0502    //The application works on WXP SP2, Windows Sever 2003 or higher platforms.
#undef    _WIN32_IE         // See sdkddkver.h
#define   _WIN32_IE                                 0x0600    //The application requires the Internet Explorer 6.0.

//Define WIN32_LEAN_AND_MEAN to exclude APIs such as Cryptography, DDE, RPC, Shell, and Windows Sockets. 
//#define WIN32_LEAN_AND_MEAN  

#define _CRT_SECURE_NO_WARNINGS           //For eliminating the "unsafe" warnings (C4996).
#define _CRT_SECURE_NO_DEPRECATE          //For eliminating the deprecation warnings.
#define _CRT_NON_CONFORMING_SWPRINTFS     //In Visual C++ 2005, swprintf conforms to the ISO C Standard, which requires the second parameter, 
                                          //count, of type size_t. This definition forces the old nonstandard behavior.

#include "Macros.h"
#include "CppStringsStreams.h"

#pragma   warning(push, 3)	//If the highest diagnostic level is set for windows.h header file, the file itself brings 
                            //lots of insignificant warnings, primarily because of nonstandard C language constructions, used in it.
#include  <WINDOWS.H>       //Header file for Windows functions and services.
#pragma   warning(pop)
#pragma   warning(push, 4)  //Setting the highest diagnostic level during the compilation.

#include  <WINDOWSX.H>      //Macro APIs, window message crackers, and control APIs.

#include  <WINERROR.H>      //Windows error codes.

//Common control library
#include  <COMMCTRL.H>
#pragma comment (lib, "comctl32.lib")   //The library, necessary for functionality of the Common control library.

//For SetupDiXxx() functions.
#include  <SETUPAPI.H>
#pragma comment (lib, "setupapi.lib")   //The library, necessary for compiling with the SetupDiXxx() fucntions.

//The header is necessary for working with GUIDs in USER-MODE and for COM.
#include  <OBJBASE.H>

//The file for registry strings definitions.
#include  <REGSTR.H>


//--------------------------- SOME USEFUL MACROS ---------------------------------

//The size of system's memory page (in bytes).
#define   PAGE_SIZE                                 4096    //ATTENTION! The size may vary depending on the OS and HW platform.

//The size of the Level 1 processor cache (in bytes).
#define   L1_CACHE_LENGTH                           32      //For x86 platforms, the cache line is 32 bytes.
//For Alpha platforms, the cache line is 64 bytes.

//Macros for analyzing the Windows status codes.
#define   IS_STATUS_CODE_ERROR(errorCode)           ((errorCode)&0x80000000 && (errorCode)&0x40000000)
#define   IS_STATUS_CODE_SUCCESS(errorCode)         (!((errorCode)&0x80000000) && !((errorCode)&0x40000000))
#define   IS_STATUS_CODE_INFORMATIONAL(errorCode)   (!((errorCode)&0x80000000) && (errorCode)&0x40000000)
#define   IS_STATUS_CODE_WARNING(errorCode)         ((errorCode)&0x80000000 && !((errorCode)&0x40000000))

//The macro, which makes easier porting a code from using Win32 API CreateThread() function
//to using _beginthreadex() C/C++ Run-Time Library function.
//For the _beginthreadex() function inclusion of <process.h> file is necessary.
typedef   unsigned(__stdcall *PTHREAD_START_FUNCTION)(void *);
#define   _BEGINTHREADEX_WRAPPER(lpThreadAttributes, dwStackSize, lpStartAddress, lpParameter, dwCreationFlags, lpThreadId) \
                                (HANDLE)_beginthreadex( \
                                                        (void*) lpThreadAttributes, \
                                                        (unsigned) dwStackSize, \
                                                        (PTHREAD_START_FUNCTION)lpStartAddress, \
                                                        (void*)lpParameter, \
                                                        (unsigned) dwCreationFlags, \
                                                        (unsigned*)lpThreadId \
                                                        )

///////////////////////////// SOME USEFUL CLASSES /////////////////////////////////
#include <vector>
#include <map>
#include <utility>  // for std::pair, std::make_pair()
namespace Utils
{
    class CommandLineOfProcess
    {
        std::vector<std::wstring>   wParsedCmdLnParams;
        std::vector<std::string>    aParsedCmdLnParams;

    public:
        CommandLineOfProcess()
        {
            std::wstring wstrFullCmdLine{ ::GetCommandLineW() };
            LPWSTR *wszArgsList;
            int nArgs;
            wszArgsList = ::CommandLineToArgvW(wstrFullCmdLine.c_str(), &nArgs);
            for (int i = 0; i < nArgs; ++i)
            {
                wParsedCmdLnParams.push_back(std::wstring{ wszArgsList[i] });
                CHAR cBuf[1024]{ '\0' };
                wcstombs(cBuf, wszArgsList[i], ELEMENTS_IN_ARRAY(cBuf)-1);
                aParsedCmdLnParams.push_back(std::string{ cBuf });
            }
            LocalFree(wszArgsList);
        }

        CommandLineOfProcess(int argc, TCHAR* argv[])
        {
            for (int i = 0; i < argc; ++i)
            {
                #ifdef _UNICODE
                wParsedCmdLnParams.push_back(std::wstring{ argv[i] });
                CHAR cBuf[1024]{ '\0' };
                wcstombs(cBuf, argv[i], ELEMENTS_IN_ARRAY(cBuf)-1);
                aParsedCmdLnParams.push_back(std::string{ cBuf });
                #else
                aParsedCmdLnParams.push_back(std::string{ argv[i] });
                WCHAR wBuf[1024]{ L'\0' };
                mbstowcs(wBuf, argv[i], ELEMENTS_IN_ARRAY(wBuf)-1);
                wParsedCmdLnParams.push_back(std::wstring{ wBuf });
                #endif
            }
        }

        CommandLineOfProcess(const CommandLineOfProcess&) = delete;
        CommandLineOfProcess& operator=(const CommandLineOfProcess&) = delete;
        CommandLineOfProcess(CommandLineOfProcess&&) = delete;
        CommandLineOfProcess& operator=(CommandLineOfProcess&&) = delete;

        const std::vector<std::wstring>& CommandLineParamsW() const
        {
            return wParsedCmdLnParams;
        }

        const std::vector<std::string>& CommandLineParamsA() const
        {
            return aParsedCmdLnParams;
        }

        #ifdef _UNICODE
        #define CommandLineParams CommandLineParamsW
        #else
        #define CommandLineParams CommandLineParamsA
        #endif
    };

    //---------------------------------------------------------------------------------
    class EnvVarsOfProcess
    {
        // envVarName, envVarValue pair
        typedef std::pair<std::wstring, std::wstring>   wEnvVarPair;
        typedef std::pair<std::string, std::string>     aEnvVarPair;
        #ifdef _UNICODE
        #define _tEnvVarPair wEnvVarPair
        #else
        #define _tEnvVarPair aEnvVarPair
        #endif
        std::map<std::wstring, std::wstring>    wEnvVars;
        std::map<std::string, std::string>      aEnvVars;

    public:
        EnvVarsOfProcess()
        {
            LPTCH envStrings = ::GetEnvironmentStrings();
            LPTCH currEnvStringPtr = envStrings;
            size_t currEnvStrLen;
            while ( (currEnvStrLen = _tcslen(currEnvStringPtr)) > 0 )
            {
                AddEnvVar(currEnvStringPtr);
                currEnvStringPtr += currEnvStrLen + 1;
            }
            ::FreeEnvironmentStrings(envStrings);
        }

        EnvVarsOfProcess(TCHAR* envp[])
        {
            int i = 0;
            while (envp[i])
                AddEnvVar(envp[i++]);
        }

        EnvVarsOfProcess(const EnvVarsOfProcess&) = delete;
        EnvVarsOfProcess& operator=(const EnvVarsOfProcess&) = delete;
        EnvVarsOfProcess(EnvVarsOfProcess&&) = delete;
        EnvVarsOfProcess& operator=(EnvVarsOfProcess&&) = delete;

        const std::map<std::wstring, std::wstring>& EnvVarsW() const
        {
            return wEnvVars;
        }

        const std::map<std::string, std::string>& EnvVarsA() const
        {
            return aEnvVars;
        }

        #ifdef _UNICODE
        #define EnvVars     EnvVarsW
        #else
        #define EnvVars     EnvVarsA
        #endif

    private:
        void AddEnvVar(PTSTR unparsedEnvVar)
        {
            std::_tstring _tstrWholeEnvVar{ unparsedEnvVar };
            std::_tstring::size_type equalSignPos = _tstrWholeEnvVar.find_last_of(_TEXT('='));
            std::_tstring _tstrEnvVarName{ _tstrWholeEnvVar, 0, equalSignPos };
            std::_tstring _tstrEnvVarValue{ _tstrWholeEnvVar, equalSignPos + 1 };
            #ifdef UNICODE
            wEnvVars.insert(std::make_pair(static_cast<std::wstring>(_tstrEnvVarName), static_cast<std::wstring>(_tstrEnvVarValue)));
            CHAR bufForName[256];
            CHAR bufForValue[4096];
            wcstombs(bufForName, _tstrEnvVarName.c_str(), ELEMENTS_IN_ARRAY(bufForName)-1);
            wcstombs(bufForValue, _tstrEnvVarValue.c_str(), ELEMENTS_IN_ARRAY(bufForValue)-1);
            aEnvVars.insert(std::make_pair(std::string{ bufForName }, std::string{ bufForValue }));
            #else
            aEnvVars.insert(std::make_pair(static_cast<std::string>(_tstrEnvVarName), static_cast<std::string>(_tstrEnvVarValue)));
            WCHAR bufForName[256];
            WCHAR bufForValue[4096];
            mbstowcs(bufForName, _tstrEnvVarName.c_str(), ELEMENTS_IN_ARRAY(bufForName) - 1);
            mbstowcs(bufForValue, _tstrEnvVarValue.c_str(), ELEMENTS_IN_ARRAY(bufForValue) - 1);
            wEnvVars.insert(std::make_pair(std::wstring{ bufForName }, std::wstring{ bufForValue }));
            #endif
        }
    };
}