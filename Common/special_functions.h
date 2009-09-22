#ifndef __UNIOFFICE_SPECIAL_FUNCTIONS_H__
#define __UNIOFFICE_SPECIAL_FUNCTIONS_H__

#include <ole2.h>

bool Is_Variant_Null( VARIANT );

HRESULT CorrectArg(
        VARIANT value,
        VARIANT *retval);

HRESULT MakeURLFromFilename(
        BSTR value,
        BSTR *retval);

#endif //__UNIOFFICE_SPECIAL_FUNCTIONS_H__
