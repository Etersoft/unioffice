#include "../Common/special_functions.h"

#include "../Common/debug.h"

using namespace std;

bool Is_Variant_Null( VARIANT var )
{
    if ( (V_VT(&var)==VT_EMPTY) || 
         (V_VT(&var)==VT_NULL)  || 
         (V_VT(&var)==VT_ERROR) 
       ) 
        return true;
        
    return false;
}

HRESULT CorrectArg(
         VARIANT value,
         VARIANT *retval)
{
  VariantInit(retval);
  if ( V_ISBYREF( &value ) ) {  
                       
    if ( (V_VT(&value) - VT_BYREF) & VT_ARRAY ) 
    {
       V_VT(retval) = V_VT(&value) - VT_BYREF;
       V_ARRAY(retval) =*(V_ARRAYREF(&value));   
       return S_OK;           
    }  
              
    switch( V_VT(&value) - VT_BYREF ) {                    
    case VT_EMPTY:
         {
         V_VT(retval) = VT_EMPTY;
         }
         break;
         
    case VT_NULL:
         {
         V_VT(retval) = VT_NULL;
         }
         break;
         
    case VT_I2:
         {
         V_VT(retval) = VT_I2;
         V_I2(retval) = *(V_I2REF(&value));
         }
         break;
         
    case VT_I4:
         {
         V_VT(retval) = VT_I4;
         V_I4(retval) = *(V_I4REF(&value));
         }
         break;
         
    case VT_I8:
         {
         V_VT(retval) = VT_I8;
         V_I8(retval) = *(V_I4REF(&value));
         }
         break;
         
    case VT_R4:
         {
         V_VT(retval) = VT_R4;
         V_R4(retval) = *(V_R4REF(&value));
         }
         break;
         
    case VT_R8:
         {
         V_VT(retval) = VT_R8;
         V_R8(retval) = *(V_R8REF(&value));
         }
         break;
         
    case VT_CY:
         {
         V_VT(retval) = VT_CY;
         V_CY(retval) = *(V_CYREF(&value));
         }
         break;
         
    case VT_DATE:
         {
         V_VT(retval) = VT_DATE;
         V_DATE(retval) = *(V_DATEREF(&value));
         }
         break;
         
    case VT_BSTR:
         {
         V_VT(retval) = VT_BSTR;
         V_BSTR(retval) = *(V_BSTRREF(&value));
         }
         break;
         
    case VT_DISPATCH:
         {
         V_VT(retval) = VT_DISPATCH;
         V_DISPATCH(retval) = *(V_DISPATCHREF(&value));
         }
         break;
         
    case VT_BOOL:
         {
         V_VT(retval) = VT_BOOL;
         V_BOOL(retval) = *(V_BOOLREF(&value));
         }
         break;
         
    case VT_VARIANT:
         {
//        V_VT(retval) = V_VT(V_VARIANTREF(&value));
//        V_DISPATCH(retval) = V_DISPATCH(V_VARIANTREF(&value));

         *retval = *(V_VARIANTREF(&value));
         }
         break;
         
    case VT_UNKNOWN:
         {
         VariantCopy((VARIANT*)V_UNKNOWNREF(&value),retval);
/*        V_VT(retval) = VT_UNKNOWN;
        V_UNKNOWN(retval) = *(V_UNKNOWNREF(&value));*/
         }
         break;
         
    case VT_UI1:
         {
         V_VT(retval) = VT_UI1;
         V_UI1(retval) = *(V_UI1REF(&value));
         }
         break;
         
    case VT_ERROR:
         {
         V_VT(retval) = VT_ERROR;
         }
         break;
         
/*    case VT_ARRAY:{
        V_VT(retval) = VT_ARRAY;
        V_ARRAY(retval) =*(V_ARRAYREF(&value));
        break;
        }*/
        
    }
  } else {
    *retval = value; 
  }
  
  return S_OK;
}

WCHAR* insert(WCHAR* src,WCHAR* dst,unsigned int index)
{
    WCHAR* res;
    int len;
    unsigned int i,j,k;

    len = lstrlenW(src);
    len = len + lstrlenW(dst) + 1;
    res = (WCHAR*) malloc(sizeof(WCHAR)*len);
    j=0;
    for (i=0;i<lstrlenW(src);i++) {
        if (i==index) {
            for (k=0;k<lstrlenW(dst);k++) {
                *(res+j)=*(dst+k);
                j++;
            }
        }
        *(res+j)=*(src+i);
        j++;
    }
    *(res+len-1)=0;

    return res;
}

int strcmpnW(WCHAR *str1, WCHAR *str2)
{
    int i=0;
    
    while (*(str2+i)!=0) {
        if (*(str2+i)!=*(str1+i)) return 0;
        i++;
    }
    return 1;
}

HRESULT MakeURLFromFilename(
         BSTR value,
         BSTR *retval)
{   
    int i;
    WCHAR *ptr;
    WCHAR *tmp1,tmp2[] = {'2','0',0};
    WCHAR file_str[] = {'f','i','l','e',':','/','/','l','o','c','a','l','h','o','s','t','/',0};
    WCHAR http[] = {'h','t','t','p',0};
    WCHAR https[] = {'h','t','t','p','s',0};
    WCHAR ftp[] = {'f','t','p',0};

    ptr = SysAllocString( value );
    if ((strcmpnW(ptr, http)+strcmpnW(ptr, https)+strcmpnW(ptr, ftp))==0) {
        i=0;
        while (*(ptr+i)!=0) {
            if (*(ptr+i)==' ') {
                *(ptr+i)='%';
            tmp1=insert(ptr,tmp2,i+1);
            ptr = tmp1;
        }
        if (*(ptr+i) == '\\')
            *(ptr+i) = '/';
        i++;
        }
        tmp1=insert(ptr,file_str,0);
        ptr = tmp1;
        }

    *retval = ptr;
 
    return S_OK;
}



