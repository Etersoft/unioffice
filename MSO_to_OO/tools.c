
#define COBJMACROS

#include "tools.h"
#include "debug.h"
#include <stdio.h>

/*
The purpose of the AutoWrap() function in this sample is to wrap the calls for GetIDsOfNames and Invoke 
to facilitate automation with straight C++. This function described in MSDN KB 238393.
*/
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs, ...)
{
    int i;
    /* Variables used...*/
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    
    /* VARIANT *pArgs;*/
    VARIANTARG vararg[10];
    EXCEPINFO exi;

    /* Begin variable-argument list...*/
    va_list marker;
    va_start(marker, cArgs);

    if (!pDisp)
       return S_FALSE;

    /* Get DISPID for name passed...*/
    hr = IDispatch_GetIDsOfNames(pDisp,&IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
       WTRACE(L"Object doesn`t have method ---> %s \n",ptName);
       return hr;
    }
    /* Extract arguments...
    Allocate memory for arguments...*/
    for (i=0; i<cArgs; i++) {
       VariantInit(&vararg[i]);
       vararg[i] = va_arg(marker, VARIANT);
    }

    /* Build DISPPARAMS*/
    dp.cArgs = cArgs;
    dp.rgvarg = vararg;

    /* Handle special-case for property-puts!*/
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    /* Make the call!*/
    hr = IDispatch_Invoke(pDisp,dispID, &IID_NULL, /*LOCALE_SYSTEM_DEFAULT*/ LOCALE_USER_DEFAULT, autoType, &dp, pvResult, &exi, NULL);
    if (FAILED(hr)) {
        switch(hr) {
        case DISP_E_BADPARAMCOUNT: TRACE("DISP_E_BADPARAMCOUNT \n");break;
        case DISP_E_BADVARTYPE: TRACE("DISP_E_BADVARTYPE \n");break;
        case DISP_E_EXCEPTION: TRACE("DISP_E_EXCEPTION \n");break;
        case DISP_E_MEMBERNOTFOUND: TRACE("DISP_E_MEMBERNOTFOUND \n");break;
        case DISP_E_NONAMEDARGS: TRACE("DISP_E_NONAMEDARGS \n");break;
        case DISP_E_OVERFLOW: TRACE("DISP_E_OVERFLOW \n");break;
        case DISP_E_PARAMNOTFOUND: TRACE("DISP_E_PARAMNOTFOUND \n");break;
        case DISP_E_TYPEMISMATCH: TRACE("DISP_E_TYPEMISMATCH \n");break;
        case DISP_E_UNKNOWNINTERFACE: TRACE("DISP_E_UNKNOWNINTERFACE \n");break;
        case DISP_E_UNKNOWNLCID: TRACE("DISP_E_UNKNOWNLCID \n");break;
        case DISP_E_PARAMNOTOPTIONAL: TRACE("DISP_E_PARAMNOTOPTIONAL \n");break;
        }

        WTRACE(L"Error in method %s -------> %s \n",ptName,exi.bstrDescription);
        return hr;
    }
    /* End variable-argument section...*/
    va_end(marker);

    return hr;
}
