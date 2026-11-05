/* Run only in a disposable Wine prefix with both UniOffice bitnesses installed. */
#include <windows.h>
#include <stdio.h>
#include <string.h>

typedef HRESULT (WINAPI *registration_fn)(void);
#define TYPELIB_KEY "TypeLib\\{00020813-0000-0000-C000-000000000046}\\1.0\\0\\"
#ifdef _WIN64
#define OWN_KEY TYPELIB_KEY "win64"
#define PEER_KEY TYPELIB_KEY "win32"
#else
#define OWN_KEY TYPELIB_KEY "win32"
#define PEER_KEY TYPELIB_KEY "win64"
#endif

static LONG registered_path(const char *key, char *value, DWORD capacity)
{
    HKEY handle;
    LONG result = RegOpenKeyExA(HKEY_CLASSES_ROOT, key, 0, KEY_READ, &handle);
    if (result) return result;
    result = RegQueryValueExA(handle, NULL, NULL, NULL, (BYTE *)value, &capacity);
    RegCloseKey(handle);
    return result;
}

int __cdecl main(int argc, char **argv)
{
    HMODULE module;
    registration_fn reg, unreg;
    char peer_before[1024], peer_after[1024], own[1024], progid[1024];
    LONG peer_status, own_status;
    HRESULT hr;
    int failures = 0;
    if (argc != 2) return 2;
    module = LoadLibraryA(argv[1]);
    if (!module) return 2;
    reg = (registration_fn)GetProcAddress(module, "DllRegisterServer");
    unreg = (registration_fn)GetProcAddress(module, "DllUnregisterServer");
    if (!reg || !unreg) { FreeLibrary(module); return 2; }
    peer_status = registered_path(PEER_KEY, peer_before, sizeof(peer_before));
    if (peer_status != ERROR_SUCCESS) {
        fprintf(stderr, "Peer typelib must be registered before this test\n");
        FreeLibrary(module);
        return 2;
    }
    hr = reg();
    if (FAILED(hr)) { FreeLibrary(module); return 2; }
    if (registered_path(OWN_KEY, own, sizeof(own)) != ERROR_SUCCESS) ++failures;
    if (registered_path("Excel.Application\\CLSID", progid, sizeof(progid)) != ERROR_SUCCESS) ++failures;
    hr = unreg();
    if (FAILED(hr)) ++failures;
    own_status = registered_path(OWN_KEY, own, sizeof(own));
    if (own_status != ERROR_FILE_NOT_FOUND && own_status != ERROR_PATH_NOT_FOUND) ++failures;
    peer_status = registered_path(PEER_KEY, peer_after, sizeof(peer_after));
    if (peer_status != ERROR_SUCCESS || strcmp(peer_before, peer_after)) ++failures;
    if (registered_path("Excel.Application\\CLSID", progid, sizeof(progid)) != ERROR_SUCCESS) ++failures;
    /* Restore own registration even if assertions failed. */
    hr = reg();
    if (FAILED(hr)) ++failures;
    FreeLibrary(module);
    printf("TypeLib registration (%u-bit): %d failure(s)\n", (unsigned)(sizeof(void *) * 8), failures);
    return failures ? 1 : 0;
}
