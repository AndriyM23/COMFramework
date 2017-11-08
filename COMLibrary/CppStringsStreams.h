#pragma once

// Support Unicode/Ansi compatibility within a single source.

//Controlling the program's character-set usage.
//#undef    UNICODE     //If this string is commented, UNICODE/ANSI compilation of the program
//can only be controlled via project's settings.
//Otherwise, if this string is uncommented, and this header is the first one
//within all project's headers inclusions, then this file's UNICODE/ANSI 
//compilation settings prevail.
//#define   UNICODE     //This preprocessor variable defintion causes the program be compiled for compliance with Unicode.
//Comment this definition and rebuild your project to make it ANSI-compatible.

//Before the inclusion of standard Windows headers, the Unicode/ANSI settings must be done properly. The following #if
//expression adjusts the settings automatically. It's necessary, that either both preprocessor variables 
//_UNICODE and UNICODE are defined or none. This is what these preprocessor variables are for:
//  - UNICODE	    -	preprocessor variable for WinAPI functions.
//  - _UNICODE	  -	preprocessor variable for C/C++ Library functions.
#if (defined(UNICODE) && !defined(_UNICODE))
#define _UNICODE
#pragma message("UNICODE-was defined, _UNICODE-was undefined. Defined: _UNICODE")
#elif (!defined(UNICODE) && defined(_UNICODE))
#define UNICODE
#pragma message("UNICODE-was undefined, _UNICODE-was defined. Defined: UNICODE")
#elif (defined(UNICODE) && defined(_UNICODE))
#pragma message("UNICODE and _UNICODE defined.")
#else
#pragma message("UNICODE and _UNICODE undefined.")
#endif

#include <wchar.h>
#include <tchar.h>      // For _vstprintf()-like macroses and other _txxx definitions.
#include <stdlib.h>     // For wcstombs()-like functions.
#include <string>
#include <iostream>
#include <sstream>
#include <fstream>
#include <iomanip>

namespace std
{

#ifdef _UNICODE

    typedef std::wstring         _tstring;         //Equivalent is: typedef basic_string<wchar_t>                                _tstring;

    typedef std::wistream        _tistream;        //Equivalent is: typedef basic_istream<wchar_t, char_traits<wchar_t> >        _tistream;
    typedef std::wostream        _tostream;        //Equivalent is: typedef basic_ostream<wchar_t, char_traits<wchar_t> >        _tostream;
    typedef std::wiostream       _tiostream;       //Equivalent is: typedef basic_iostream<wchar_t, char_traits<wchar_t> >       _tiostream;

    typedef std::wistringstream  _tistringstream;  //Equivalent is: typedef basic_istringstream<wchar_t, char_traits<wchar_t> >  _tistringstream;
    typedef std::wostringstream  _tostringstream;  //Equivalent is: typedef basic_ostringstream<wchar_t, char_traits<wchar_t> >  _tostringstream;
    typedef std::wstringstream   _tstringstream;   //Equivalent is: typedef basic_stringstream<wchar_t, char_traits<wchar_t> >   _tstringstream;

    typedef std::wifstream       _tifstream;       //Equivalent is: typedef basic_ifstream<wchar_t, char_traits<wchar_t> >       _tifstream;
    typedef std::wofstream       _tofstream;       //Equivalent is: typedef basic_ofstream<wchar_t, char_traits<wchar_t> >       _tofstream;
    typedef std::wfstream        _tfstream;        //Equivalent is: typedef basic_fstream<wchar_t, char_traits<wchar_t> >        _tfstream;

    //Defining some alternative preprocessor names of global variables for UNICODE/ANSI source-code compatibility.
    #define _tcin                std::wcin
    #define _tcout               std::wcout
    #define _tcerr               std::wcerr

#else

    typedef std::string          _tstring;         //Equivalent is: typedef basic_string<char>                             _tstring;

    typedef std::istream         _tistream;        //Equivalent is: typedef basic_istream<char, char_traits<char> >        _tistream;
    typedef std::ostream         _tostream;        //Equivalent is: typedef basic_ostream<char, char_traits<char> >        _tostream;
    typedef std::iostream        _tiostream;       //Equivalent is: typedef basic_iostream<char, char_traits<char> >       _tiostream;

    typedef std::istringstream   _tistringstream;  //Equivalent is: typedef basic_istringstream<char, char_traits<char> >  _tistringstream;
    typedef std::ostringstream   _tostringstream;  //Equivalent is: typedef basic_ostringstream<char, char_traits<char> >  _tostringstream;
    typedef std::stringstream    _tstringstream;   //Equivalent is: typedef basic_stringstream<char, char_traits<char> >   _tstringstream;

    typedef std::ifstream        _tifstream;       //Equivalent is: typedef basic_ifstream<char, char_traits<char> >       _tifstream;
    typedef std::ofstream        _tofstream;       //Equivalent is: typedef basic_ofstream<char, char_traits<char> >       _tofstream;
    typedef std::fstream         _tfstream;        //Equivalent is: typedef basic_fstream<char, char_traits<char> >        _tfstream;

    //Defining some alternative preprocessor names of global variables for UNICODE/ANSI source-code compatibility.
    #define _tcin                std::cin
    #define _tcout               std::cout
    #define _tcerr               std::cerr

#endif

// Some convenient functions.

#define TRIMSTR_TRIM_FROM_BEGIN                             ((unsigned long)0x00000001UL)
#define TRIMSTR_TRIM_FROM_END                               ((unsigned long)0x00000002UL)
#define TRIMSTR_TREAT_TOBEFOUNDSTRING_AS_A_WHOLE            ((unsigned long)0x00000004UL)
#define TRIMSTR_IGNORECASE                                  ((unsigned long)0x00000008UL)
#define TRIMSTR_MAX_ITERATIONS                              ((unsigned long)0xFFFFFFFFUL)
_tstring TrimString ( const _tstring& str, 
                      const _tstring& stringOrCharsToBeFound = std::_tstring{ _TEXT(" ") },
                      unsigned long flags = TRIMSTR_TRIM_FROM_BEGIN | TRIMSTR_TRIM_FROM_END, 
                      unsigned long nIterations = 1
                    );

}  //namespace std
