#pragma     once    //Specifies that the file will be included (opened) only once by the compiler when compiling a source code file.

//Macros for lexem stringizing.
#define   _STRINGIZE_LEXEM(lexem)                   #lexem
#define   STRINGIZE_LEXEM(lexem)                    _STRINGIZE_LEXEM(lexem)   //Always use this macro!!!

//Macros for lexem charizing (Microsoft specific).
#define   _CHARIZE_LEXEM(lexem)                     #@lexem
#define   CHARIZE_LEXEM(lexem)                      _CHARIZE_LEXEM(lexem)     //Always use this macro!!!

//Macros for conversion of ANSI string constants into Unicode string constants.
#define   _TEXT_TO_UNICODE(text)                    L##text
#define   TEXT_TO_UNICODE(text)                     _TEXT_TO_UNICODE(text)    //Always use this macro!!!

//Macros for conversion the lexems into Unicode string constants.
#define   _LEXEM_TO_UNICODE(lexem)                  L## #lexem
#define   LEXEM_TO_UNICODE(lexem)                   _LEXEM_TO_UNICODE(lexem)  //Always use this macro!!!

//Macro for structures padding. Instead of this macro, the Microsoft-specific modifier __declspec(align(#)) can be used like this:
//__declspec(align(PadBoundary)) TYPE AlignedElement;
#define   PAD_STRUCTURE(PadderName, WhereStartPadding, PadBoundary)       \
                    BYTE PadderName [(PadBoundary)-(WhereStartPadding)%(PadBoundary)]

//Macro identifies the amount of elements in an array.
//The array must be declared as an array (according to the C language requirements).
#define   ELEMENTS_IN_ARRAY(arrayName)              (sizeof(arrayName)/sizeof(arrayName[0]))
