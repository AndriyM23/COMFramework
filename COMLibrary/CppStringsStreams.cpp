#include "CppStringsStreams.h"
#include <algorithm>    // for std:transform
#include <cctype>       // for std:tolower

namespace std
{
    _tstring TrimString(const _tstring& str,
                        const _tstring& stringOrCharsToBeFound,
                        unsigned long flags,
                        unsigned long nIterations )
    {
        _tstring resultString{ str };
        if (str.empty() || stringOrCharsToBeFound.empty() || nIterations == 0 ||
            (flags & (TRIMSTR_TRIM_FROM_BEGIN | TRIMSTR_TRIM_FROM_END)) == 0 ||
            ((flags & TRIMSTR_TREAT_TOBEFOUNDSTRING_AS_A_WHOLE) && str.size() < stringOrCharsToBeFound.size()) )
            return std::move(resultString);

        _tstring stringOrCharsToBeFound_local{stringOrCharsToBeFound};
        if (flags & TRIMSTR_IGNORECASE)
            transform(stringOrCharsToBeFound_local.begin(), stringOrCharsToBeFound_local.end(), stringOrCharsToBeFound_local.begin(), (int(*) (int))_totlower);
        bool searchFromBeginning = (flags & TRIMSTR_TRIM_FROM_BEGIN) != 0;
        bool searchFromEnd = (flags & TRIMSTR_TRIM_FROM_END) != 0;
        
        if (flags & TRIMSTR_TREAT_TOBEFOUNDSTRING_AS_A_WHOLE)
            for (unsigned long i = 0; i<nIterations && (searchFromBeginning || searchFromEnd); ++i)
            {
                if (searchFromBeginning)
                {
                    _tstring startsWithStr = resultString.substr(0, stringOrCharsToBeFound_local.size());
                    if (flags & TRIMSTR_IGNORECASE)
                        transform(startsWithStr.begin(), startsWithStr.end(), startsWithStr.begin(), (int(*) (int))_totlower);

                    if (startsWithStr == stringOrCharsToBeFound_local)
                        resultString.erase(0, stringOrCharsToBeFound_local.size());
                    else
                        searchFromBeginning = false;
                }

                if (searchFromEnd)
                {
                    _tstring endsWithStr = resultString.substr(resultString.size() - stringOrCharsToBeFound_local.size());
                    if (flags & TRIMSTR_IGNORECASE)
                        transform(endsWithStr.begin(), endsWithStr.end(), endsWithStr.begin(), (int(*) (int))_totlower);

                    if (endsWithStr == stringOrCharsToBeFound_local)
                        resultString.erase(resultString.size() - stringOrCharsToBeFound_local.size());
                    else
                        searchFromEnd = false;
                }
            }
        else
        {
            _tstring resultString_local{ str };
            if (flags & TRIMSTR_IGNORECASE)
                transform(resultString_local.begin(), resultString_local.end(), resultString_local.begin(), (int(*) (int))_totlower);
            for (unsigned long i = 0; i<nIterations && (searchFromBeginning || searchFromEnd); ++i)
            {
                if (searchFromBeginning)
                {
                    _tstring::size_type begSymPos = resultString_local.find_first_of(stringOrCharsToBeFound_local);
                    if (begSymPos == _tstring::npos)
                        searchFromBeginning = searchFromEnd = false;
                    else if (begSymPos > 0)
                        searchFromBeginning = false;
                    else
                    {
                        resultString.erase(0, 1);
                        resultString_local.erase(0, 1);
                    }
                }

                if (searchFromEnd)
                {
                    _tstring::size_type endSymPos = resultString_local.find_last_of(stringOrCharsToBeFound_local);
                    if (endSymPos == _tstring::npos)
                        searchFromBeginning = searchFromEnd = false;
                    else if (endSymPos < resultString.size()-1)
                        searchFromEnd = false;
                    else
                    {
                        resultString.erase(endSymPos);
                        resultString_local.erase(endSymPos);
                    }
                }
            }
        }

        return std::move(resultString);
    }
}  //namespace std
