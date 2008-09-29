/*
 * Register and unregister ooffice.dll functions
 *
 * Copyright (C) 2003 John K. Hohm
 * Copyright (C) 2007 Roy Shea (Google)
 * Copyright (C) 2008 Sinitsin Ivan (Etersoft) <ivan@etersoft.ru>
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA
 */

#include "unioffice_word_private.h"
#include <shlwapi.h>

/*
 * Interface for self-registering
 */

struct regsvr_interface
{
    IID const *iid;             /* NULL for end of list */
    LPCSTR name;                /* can be NULL to omit */
    IID const *base_iid;        /* can be NULL to omit */
    int num_methods;            /* can be <0 to omit */
    CLSID const *ps_clsid;      /* can be NULL to omit */
    CLSID const *ps_clsid32;    /* can be NULL to omit */
};

static HRESULT register_interfaces(struct regsvr_interface const *list);
static HRESULT unregister_interfaces(struct regsvr_interface const *list);

struct regsvr_coclass
{
    CLSID const *clsid;         /* NULL for end of list */
    LPCSTR name;                /* can be NULL to omit */
    LPCSTR ips;                 /* can be NULL to omit */
    LPCSTR ips32;               /* can be NULL to omit */
    LPCSTR ips32_tmodel;        /* can be NULL to omit */
};

static HRESULT register_coclasses(struct regsvr_coclass const *list);
static HRESULT unregister_coclasses(struct regsvr_coclass const *list);

static HRESULT register_classes();
static HRESULT unregister_classes();

/*
 * Static string constants
 */
static WCHAR const interface_keyname[10] = {
    'I', 'n', 't', 'e', 'r', 'f', 'a', 'c', 'e', 0 };
static WCHAR const base_ifa_keyname[14] = {
    'B', 'a', 's', 'e', 'I', 'n', 't', 'e', 'r', 'f', 'a', 'c', 'e', 0 };
static WCHAR const num_methods_keyname[11] = {
    'N', 'u', 'm', 'M', 'e', 't', 'h', 'o', 'd', 's', 0 };
static WCHAR const ps_clsid_keyname[15] = {
    'P', 'r', 'o', 'x', 'y', 'S', 't', 'u', 'b', 'C', 'l', 's', 'i', 'd', 0 };
static WCHAR const ps_clsid32_keyname[17] = {
    'P', 'r', 'o', 'x', 'y', 'S', 't', 'u', 'b',
    'C', 'l', 's', 'i', 'd', '3', '2', 0 };
static WCHAR const clsid_keyname[6] = {
    'C', 'L', 'S', 'I', 'D', 0 };
static WCHAR const ips_keyname[13] = {
    'I', 'n', 'P', 'r', 'o', 'c', 'S', 'e', 'r', 'v', 'e', 'r', 0 };
static WCHAR const ips32_keyname[15] = {
    'I', 'n', 'P', 'r', 'o', 'c', 'S', 'e', 'r', 'v', 'e', 'r', '3', '2', 0 };
static char const tmodel_valuename[] = "ThreadingModel";


static WCHAR class_name[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n',0};
static WCHAR class_name1[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1',0};
static WCHAR class_name2[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','2',0};
static WCHAR class_name3[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','3',0};
static WCHAR class_name4[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','4',0};
static WCHAR class_name5[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','5',0};
static WCHAR class_name6[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','6',0};
static WCHAR class_name7[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','7',0};
static WCHAR class_name8[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','8',0};
static WCHAR class_name9[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','9',0};
static WCHAR class_name10[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1','0',0};
static WCHAR class_name11[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1','1',0};

static WCHAR const class_CLSID[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','\\','C','L','S','I','D',0};
static WCHAR const class_CurVer[] = {  
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','\\','C','u','r','V','e','r',0};
static WCHAR const str_curver[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1',0};

static WCHAR const class_CLSID1[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID2[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','2','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID3[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','3','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID4[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','4','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID5[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','5','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID6[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','6','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID7[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','7','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID8[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','8','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID9[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','9','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID10[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1','0','\\','C','L','S','I','D',0};
static WCHAR const class_CLSID11[] = {
    'W','o','r','d','.','A','p','p','l','i','c','a','t','i','o','n','.','1','1','\\','C','L','S','I','D',0};


/*
 * Static helper functions
 */
static LONG register_key_guid(HKEY base, WCHAR const *name, GUID const *guid);
static LONG register_key_defvalueW(HKEY base, WCHAR const *name,
        WCHAR const *value);
static LONG register_key_defvalueA(HKEY base, WCHAR const *name,
        char const *value);
static LONG recursive_delete_key(HKEY key, WCHAR *key_name);


/*
 * Register_interfaces
 */
static HRESULT register_interfaces(struct regsvr_interface const *list)
{
    LONG res = ERROR_SUCCESS;
    HKEY interface_key;

    res = RegCreateKeyExW(HKEY_CLASSES_ROOT, interface_keyname, 0, NULL, 0,
            KEY_READ | KEY_WRITE, NULL, &interface_key, NULL);
    if (res != ERROR_SUCCESS) goto error_return;

    for (; res == ERROR_SUCCESS && list->iid; ++list) {
        WCHAR buf[39];
        HKEY iid_key;

        StringFromGUID2(*(list->iid), buf, 39);
        res = RegCreateKeyExW(interface_key, buf, 0, NULL, 0,
                KEY_READ | KEY_WRITE, NULL, &iid_key, NULL);
        if (res != ERROR_SUCCESS) goto error_close_interface_key;

        if (list->name) {
            res = RegSetValueExA(iid_key, NULL, 0, REG_SZ,
                    (CONST BYTE*)(list->name),
                    strlen(list->name) + 1);
            if (res != ERROR_SUCCESS) goto error_close_iid_key;
        }

        if (list->base_iid) {
            res = register_key_guid(iid_key, base_ifa_keyname, list->base_iid);
            if (res != ERROR_SUCCESS) goto error_close_iid_key;
        }

        if (0 <= list->num_methods) {
            static WCHAR const fmt[3] = { '%', 'd', 0 };
            HKEY key;

            res = RegCreateKeyExW(iid_key, num_methods_keyname, 0, NULL, 0,
                    KEY_READ | KEY_WRITE, NULL, &key, NULL);
            if (res != ERROR_SUCCESS) goto error_close_iid_key;

            wsprintfW(buf, fmt, list->num_methods);
            res = RegSetValueExW(key, NULL, 0, REG_SZ,
                    (CONST BYTE*)buf,
                    (lstrlenW(buf) + 1) * sizeof(WCHAR));
            RegCloseKey(key);

            if (res != ERROR_SUCCESS) goto error_close_iid_key;
        }

        if (list->ps_clsid) {
            res = register_key_guid(iid_key, ps_clsid_keyname, list->ps_clsid);
            if (res != ERROR_SUCCESS) goto error_close_iid_key;
        }

        if (list->ps_clsid32) {
            res = register_key_guid(iid_key, ps_clsid32_keyname, list->ps_clsid32);
            if (res != ERROR_SUCCESS) goto error_close_iid_key;
        }

error_close_iid_key:
        RegCloseKey(iid_key);
    }

error_close_interface_key:
    RegCloseKey(interface_key);
error_return:
    return res != ERROR_SUCCESS ? HRESULT_FROM_WIN32(res) : S_OK;
}

/*
 * unregister_interfaces
 */
static HRESULT unregister_interfaces(struct regsvr_interface const *list)
{
    LONG res = ERROR_SUCCESS;
    HKEY interface_key;

    res = RegOpenKeyExW(HKEY_CLASSES_ROOT, interface_keyname, 0,
            KEY_READ | KEY_WRITE, &interface_key);
    if (res == ERROR_FILE_NOT_FOUND) return S_OK;
    if (res != ERROR_SUCCESS) goto error_return;

    for (; res == ERROR_SUCCESS && list->iid; ++list) {
        WCHAR buf[39];
        StringFromGUID2(*(list->iid), buf, 39);
        res = SHDeleteKeyW(interface_key, buf);
        if (res == ERROR_FILE_NOT_FOUND) res = ERROR_SUCCESS;
        if (res != ERROR_SUCCESS) goto error_close_interface_key;
    }

error_close_interface_key:
    RegCloseKey(interface_key);
error_return:
    return res != ERROR_SUCCESS ? HRESULT_FROM_WIN32(res) : S_OK;
}

/*
 * register_coclasses
 */
static HRESULT register_coclasses(struct regsvr_coclass const *list)
{
    LONG res = ERROR_SUCCESS;
    HKEY coclass_key;

    res = RegCreateKeyExW(HKEY_CLASSES_ROOT, clsid_keyname, 0, NULL, 0,
            KEY_READ | KEY_WRITE, NULL, &coclass_key, NULL);
    if (res != ERROR_SUCCESS) goto error_return;

    for (; res == ERROR_SUCCESS && list->clsid; ++list) {
        WCHAR buf[39];
        HKEY clsid_key;

        StringFromGUID2(*(list->clsid), buf, 39);
        res = RegCreateKeyExW(coclass_key, buf, 0, NULL, 0,
                KEY_READ | KEY_WRITE, NULL, &clsid_key, NULL);
        if (res != ERROR_SUCCESS) goto error_close_coclass_key;

        if (list->name) {
            res = RegSetValueExA(clsid_key, NULL, 0, REG_SZ,
                    (CONST BYTE*)(list->name),
                    strlen(list->name) + 1);
            if (res != ERROR_SUCCESS) goto error_close_clsid_key;
        }

        if (list->ips) {
            res = register_key_defvalueA(clsid_key, ips_keyname, list->ips);
            if (res != ERROR_SUCCESS) goto error_close_clsid_key;
        }

        if (list->ips32) {
            HKEY ips32_key;

            res = RegCreateKeyExW(clsid_key, ips32_keyname, 0, NULL, 0,
                    KEY_READ | KEY_WRITE, NULL,
                    &ips32_key, NULL);
            if (res != ERROR_SUCCESS) goto error_close_clsid_key;

            res = RegSetValueExA(ips32_key, NULL, 0, REG_SZ,
                    (CONST BYTE*)list->ips32,
                    lstrlenA(list->ips32) + 1);
            if (res == ERROR_SUCCESS && list->ips32_tmodel)
                res = RegSetValueExA(ips32_key, tmodel_valuename, 0, REG_SZ,
                        (CONST BYTE*)list->ips32_tmodel,
                        strlen(list->ips32_tmodel) + 1);
            RegCloseKey(ips32_key);
            if (res != ERROR_SUCCESS) goto error_close_clsid_key;
        }

error_close_clsid_key:
        RegCloseKey(clsid_key);
    }

error_close_coclass_key:
    RegCloseKey(coclass_key);
error_return:
    return res != ERROR_SUCCESS ? HRESULT_FROM_WIN32(res) : S_OK;
}

/*
 * unregister_coclasses
 */
static HRESULT unregister_coclasses(struct regsvr_coclass const *list)
{
    LONG res = ERROR_SUCCESS;
    HKEY coclass_key;

    res = RegOpenKeyExW(HKEY_CLASSES_ROOT, clsid_keyname, 0,
            KEY_READ | KEY_WRITE, &coclass_key);
    if (res == ERROR_FILE_NOT_FOUND) return S_OK;
    if (res != ERROR_SUCCESS) goto error_return;

    for (; res == ERROR_SUCCESS && list->clsid; ++list) {
        WCHAR buf[39];
        StringFromGUID2(*(list->clsid), buf, 39);
        res = SHDeleteKeyW(coclass_key, buf);
        if (res == ERROR_FILE_NOT_FOUND) res = ERROR_SUCCESS;
        if (res != ERROR_SUCCESS) goto error_close_coclass_key;
    }
    

error_close_coclass_key:
    RegCloseKey(coclass_key);
error_return:
    return res != ERROR_SUCCESS ? HRESULT_FROM_WIN32(res) : S_OK;
}


static HRESULT register_classes()
{
    HRESULT hr;
    WCHAR str_clsid[39];

    StringFromGUID2(CLSID_CApplication, str_clsid, 39);
    hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CurVer, str_curver);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID1, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID2, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID3, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID4, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID5, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID6, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID7, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID8, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID9, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID10, str_clsid);
    if (SUCCEEDED(hr))
       hr = register_key_defvalueW(HKEY_CLASSES_ROOT, class_CLSID11, str_clsid);
    return hr;
}

static HRESULT unregister_classes()
{
    HRESULT hr;
    hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name1);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name2);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name3);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name4);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name5);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name6);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name7);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name8);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name9);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name10);
    if (SUCCEEDED(hr)) 
       hr = recursive_delete_key(HKEY_CLASSES_ROOT, class_name11);

    return hr;
}

/*
 * regsvr_key_guid
 */
static LONG register_key_guid(HKEY base, WCHAR const *name, GUID const *guid)
{
    WCHAR buf[39];

    StringFromGUID2(*guid, buf, 39);
    return register_key_defvalueW(base, name, buf);
}

/*
 * regsvr_key_defvalueW
 */
static LONG register_key_defvalueW(
        HKEY base,
        WCHAR const *name,
        WCHAR const *value)
{
    LONG res;
    HKEY key;

    res = RegCreateKeyExW(base, name, 0, NULL, 0,
            KEY_READ | KEY_WRITE, NULL, &key, NULL);
    if (res != ERROR_SUCCESS) return res;
    res = RegSetValueExW(key, NULL, 0, REG_SZ, (CONST BYTE*)value,
            (lstrlenW(value) + 1) * sizeof(WCHAR));
    RegCloseKey(key);
    return res;
}

/*
 * regsvr_key_defvalueA
 */
static LONG register_key_defvalueA(
        HKEY base,
        WCHAR const *name,
        char const *value)
{
    LONG res;
    HKEY key;

    res = RegCreateKeyExW(base, name, 0, NULL, 0,
            KEY_READ | KEY_WRITE, NULL, &key, NULL);
    if (res != ERROR_SUCCESS) return res;
    res = RegSetValueExA(key, NULL, 0, REG_SZ, (CONST BYTE*)value,
            lstrlenA(value) + 1);
    RegCloseKey(key);
    return res;
}

/*
 * recursive_delete_key
 */
static LONG recursive_delete_key(HKEY key, WCHAR *key_name)
{
    LONG res;
    WCHAR subkey_name[MAX_PATH];
    DWORD cName;
    HKEY subkey;

/*    WINE_TRACE("Deleting: %s\n", wine_dbgstr_w(key_name));*/


    for (;;) {
        
        res = RegOpenKeyExW(key, key_name, 0, KEY_READ | KEY_WRITE, &subkey);
        if (res == ERROR_FILE_NOT_FOUND) {
            res = ERROR_SUCCESS;
            break;
        }
        if (res != ERROR_SUCCESS) break;
        
        cName = sizeof(subkey_name) / sizeof(WCHAR);
        res = RegEnumKeyExW(subkey, 0, subkey_name, &cName,
                NULL, NULL, NULL, NULL);
        if (res != ERROR_SUCCESS && res != ERROR_MORE_DATA) {
            res = ERROR_SUCCESS; /* presumably we're done enumerating */
            RegCloseKey(subkey);
            break;
        }

/*        WINE_TRACE("Found subkey %s/%s\n",
                wine_dbgstr_w(key_name), wine_dbgstr_w(subkey_name));*/
        res = recursive_delete_key(subkey, subkey_name);
        RegCloseKey(subkey);

/*        WINE_TRACE("Finished deleteing subkeys of %s/%s\n", 
                wine_dbgstr_w(key_name), wine_dbgstr_w(subkey_name));*/
        if (res != ERROR_SUCCESS) break;
    }

/*    WINE_TRACE("Perhaps ready to delete %s\n", wine_dbgstr_w(key_name));*/
    if (res == ERROR_SUCCESS) 
    {
        res = RegDeleteKeyW(key, key_name);
        if (res == ERROR_FILE_NOT_FOUND) {
            res = ERROR_SUCCESS;
        }
    }
/*    WINE_TRACE("Attempt to delete %s returned: 0x%08X\n", 
            wine_dbgstr_w(key_name), res);*/
    return res;
}

/*
 * Coclass list
 */
static struct regsvr_coclass const coclass_list[] = {
    {   &CLSID_CApplication,
        "Word.Application",
        NULL,
        "unioffice_word.dll",
        "Both"
    },
    { NULL }            /* list terminator */
};

/*
 * Interface list
 */
static struct regsvr_interface const interface_list[] = {
    { &IID_IApplication,
      "IApplication",
      NULL,
      13,                /*Number of methods*/
      NULL,
      NULL
    },
    { NULL }            /* list terminator */
};

#ifdef __cplusplus
extern "C" 
{
#endif

__declspec(dllexport) HRESULT __stdcall DllRegisterServer(void)
{
    HRESULT hr;

/*    TRACE("\n");*/

    hr = register_coclasses(coclass_list);
    if (SUCCEEDED(hr))
        hr = register_interfaces(interface_list);
    if (SUCCEEDED(hr))
        hr = register_classes();
    return hr;
}

__declspec(dllexport) HRESULT __stdcall DllUnregisterServer(void)
{
    HRESULT hr;

/*    TRACE("\n");*/

    hr = unregister_coclasses(coclass_list);
    if (SUCCEEDED(hr))
        hr = unregister_interfaces(interface_list);
    if (SUCCEEDED(hr))
        hr = unregister_classes();
    return hr;
}

#ifdef __cplusplus
}
#endif
