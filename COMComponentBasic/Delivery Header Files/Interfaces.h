#pragma once

#include <objbase.h>    // This inclusion makes GUIDs be declared (as external constant variables).
#include <winerror.h>

// Define success and error codes for the component
#define BASE_SUCCESS_CODE             512
#define BASE_ERROR_CODE               (BASE_SUCCESS_CODE + 100)

#define COMPONENTBASIC_S_SCCESS1      MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, (BASE_SUCCESS_CODE + 0))
#define COMPONENTBASIC_S_SCCESS2      MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, (BASE_SUCCESS_CODE + 1))
#define COMPONENTBASIC_S_SCCESS3      MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, (BASE_SUCCESS_CODE + 2))
#define COMPONENTBASIC_S_SCCESS4      MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, (BASE_SUCCESS_CODE + 3))
#define COMPONENTBASIC_S_SCCESS5      MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, (BASE_SUCCESS_CODE + 4))
#define COMPONENTBASIC_S_SCCESS6      MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, (BASE_SUCCESS_CODE + 5))

#define COMPONENTBASIC_E_ERROR1       MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, (BASE_ERROR_CODE + 0))
#define COMPONENTBASIC_E_ERROR2       MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, (BASE_ERROR_CODE + 1))
#define COMPONENTBASIC_E_ERROR3       MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, (BASE_ERROR_CODE + 2))
#define COMPONENTBASIC_E_ERROR4       MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, (BASE_ERROR_CODE + 3))
#define COMPONENTBASIC_E_ERROR5       MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, (BASE_ERROR_CODE + 4))
#define COMPONENTBASIC_E_ERROR6       MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, (BASE_ERROR_CODE + 5))

// Declare interfaces
interface IInterfBasic01 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic01() = 0;
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic01() = 0;
};
typedef IInterfBasic01  *PIInterfBasic01;

interface IInterfBasic02 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic02() = 0;
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic02() = 0;
};
typedef IInterfBasic02  *PIInterfBasic02;

interface IInterfBasic03 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic03() = 0;
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic03() = 0;
};
typedef IInterfBasic03  *PIInterfBasic03;

interface IInterfBasic04 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic04() = 0;
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic04() = 0;
};
typedef IInterfBasic04  *PIInterfBasic04;

interface IInterfBasic05 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic05() = 0;
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic05() = 0;
};
typedef IInterfBasic05  *PIInterfBasic05;

interface IInterfBasic06 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE F01_IInterfBasic06() = 0;
    virtual HRESULT STDMETHODCALLTYPE F02_IInterfBasic06() = 0;
};
typedef IInterfBasic06  *PIInterfBasic06;


#ifdef DECLARE_AND_DEFINE_GUIDS
    //If requested - we define IIDs. Done only in one source file of a component, and in one source file of a client.
    #include <initguid.h>   // This inclusion makes GUIDs be defined.
#endif

// {8FF35317-FD05-4BBF-AD7A-FA00A04EC90F}
DEFINE_GUID(IID_IInterfBasic01, 0x8ff35317, 0xfd05, 0x4bbf, 0xad, 0x7a, 0xfa, 0x0, 0xa0, 0x4e, 0xc9, 0xf);
// {3059FAED-BA24-49F9-BB97-62DAE236FBCE}
DEFINE_GUID(IID_IInterfBasic02, 0x3059faed, 0xba24, 0x49f9, 0xbb, 0x97, 0x62, 0xda, 0xe2, 0x36, 0xfb, 0xce);
// {94791F5F-E5C8-4F95-A7DB-8E91D8F04A25}
DEFINE_GUID(IID_IInterfBasic03, 0x94791f5f, 0xe5c8, 0x4f95, 0xa7, 0xdb, 0x8e, 0x91, 0xd8, 0xf0, 0x4a, 0x25);
// {9C6DF6DA-8FCE-4684-9316-CE7382DCF836}
DEFINE_GUID(IID_IInterfBasic04, 0x9c6df6da, 0x8fce, 0x4684, 0x93, 0x16, 0xce, 0x73, 0x82, 0xdc, 0xf8, 0x36);
// {04C2D129-DBAA-4A34-9D7F-5F74006E9B84}
DEFINE_GUID(IID_IInterfBasic05, 0x4c2d129, 0xdbaa, 0x4a34, 0x9d, 0x7f, 0x5f, 0x74, 0x0, 0x6e, 0x9b, 0x84);
// {BE3F1CDD-BAE4-42AF-A87D-47A7389AE234}
DEFINE_GUID(IID_IInterfBasic06, 0xbe3f1cdd, 0xbae4, 0x42af, 0xa8, 0x7d, 0x47, 0xa7, 0x38, 0x9a, 0xe2, 0x34);


// {69D4C18B-FD6B-46BE-8A71-5BEDFF73642E}
DEFINE_GUID(CLSID_COMComponent01, 0x69d4c18b, 0xfd6b, 0x46be, 0x8a, 0x71, 0x5b, 0xed, 0xff, 0x73, 0x64, 0x2e);
#define COMCOMPONENT01_FRIENDLY_NAME    "COMComponent01 (COMTrial)."
#define COMCOMPONENT01_PROGID           "COMTrial.COMComponent01.1"
#define COMCOMPONENT01_VERINDEPPROGID   "COMTrial.COMComponent01"

// {9DA6C04A-DB57-4E27-A191-29BD6256202D}
DEFINE_GUID(CLSID_COMComponent04, 0x9da6c04a, 0xdb57, 0x4e27, 0xa1, 0x91, 0x29, 0xbd, 0x62, 0x56, 0x20, 0x2d);
#define COMCOMPONENT04_FRIENDLY_NAME    "COMComponent04 (COMTrial)."
#define COMCOMPONENT04_PROGID           "COMTrial.COMComponent04.1"
#define COMCOMPONENT04_VERINDEPPROGID   "COMTrial.COMComponent04"

// {7CB7D744-621E-47BB-BD91-D8AB8DE5CEEB}
DEFINE_GUID(CLSID_COMComponentI040506, 0x7cb7d744, 0x621e, 0x47bb, 0xbd, 0x91, 0xd8, 0xab, 0x8d, 0xe5, 0xce, 0xeb);
#define COMCOMPONENTINC040506_FRIENDLY_NAME    "COMComponentI040506 (COMTrial)."
#define COMCOMPONENTINC040506_PROGID           "COMTrial.COMComponentI040506.1"
#define COMCOMPONENTINC040506_VERINDEPPROGID   "COMTrial.COMComponentI040506"

// {602396C0-0446-4D7E-96C7-E4822593EC15}
DEFINE_GUID(CLSID_COMComponentI010203A04A05A06, 0x602396c0, 0x446, 0x4d7e, 0x96, 0xc7, 0xe4, 0x82, 0x25, 0x93, 0xec, 0x15);
#define COMCOMPONENTI010203A04A05A06_FRIENDLY_NAME     "COMComponentI010203A04A05A06 (COMTrial)."
#define COMCOMPONENTI010203A04A05A06_PROGID            "COMTrial.COMComponentI010203A04A05A06.1"
#define COMCOMPONENTI010203A04A05A06_VERINDEPPROGID    "COMTrial.COMComponentI010203A04A05A06"

// {EF666192-7C0D-47DC-86B9-7F386E3F3A50}
DEFINE_GUID(CLSID_COMComponent02A01, 0xef666192, 0x7c0d, 0x47dc, 0x86, 0xb9, 0x7f, 0x38, 0x6e, 0x3f, 0x3a, 0x50);
#define COMCOMPONENT02A01_FRIENDLY_NAME     "COMComponent02A01 (COMTrial)."
#define COMCOMPONENT02A01_PROGID            "COMTrial.COMComponent02A01.1"
#define COMCOMPONENT02A01_VERINDEPPROGID    "COMTrial.COMComponent02A01"

// {74FBFBFA-8425-4669-B5B6-D341F8E06EEC}
DEFINE_GUID(CLSID_COMComponentI04A02AA01, 0x74fbfbfa, 0x8425, 0x4669, 0xb5, 0xb6, 0xd3, 0x41, 0xf8, 0xe0, 0x6e, 0xec);
#define COMCOMPONENTI04A02AA01_FRIENDLY_NAME     "COMComponentI04A02AA01 (COMTrial)."
#define COMCOMPONENTI04A02AA01_PROGID            "COMTrial.COMComponentI04A02AA01.1"
#define COMCOMPONENTI04A02AA01_VERINDEPPROGID    "COMTrial.COMComponentI04A02AA01"


// {819C70E4-41F6-462E-A983-9010EFFE42EF}
DEFINE_GUID(COMCategoryIFB01, 0x819c70e4, 0x41f6, 0x462e, 0xa9, 0x83, 0x90, 0x10, 0xef, 0xfe, 0x42, 0xef);
#define COMCATEGORYIFB01_DESCRIPTION                "COMCategoryIFB01 (COMTrial)"

// {24312B3D-DB69-43A8-A245-1FC4245611AE}
DEFINE_GUID(COMCategoryIFB04, 0x24312b3d, 0xdb69, 0x43a8, 0xa2, 0x45, 0x1f, 0xc4, 0x24, 0x56, 0x11, 0xae);
#define COMCATEGORYIFB04_DESCRIPTION                "COMCategoryIFB04 (COMTrial)"

// {24A641B0-F889-438C-9599-71EF348C8062}
DEFINE_GUID(COMCategoryIFB040506, 0x24a641b0, 0xf889, 0x438c, 0x95, 0x99, 0x71, 0xef, 0x34, 0x8c, 0x80, 0x62);
#define COMCATEGORYIFB040506_DESCRIPTION            "COMCategoryIFB040506 (COMTrial)"

// {BDAA0D97-7C54-4087-AB4B-11E03A3A578F}
DEFINE_GUID(COMCategoryIFB010203040506, 0xbdaa0d97, 0x7c54, 0x4087, 0xab, 0x4b, 0x11, 0xe0, 0x3a, 0x3a, 0x57, 0x8f);
#define COMCATEGORYIFB010203040506_DESCRIPTION      "COMCategoryIFB010203040506 (COMTrial)"


