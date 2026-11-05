/* COM lifecycle smoke test: no office installation or registry changes required. */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <stdio.h>

static const CLSID excel_clsid = {0x00024500, 0, 0, {0xc0, 0, 0, 0, 0, 0, 0, 0x46}};
static int failures;
#define CHECK(condition) do { if (!(condition)) { \
    fprintf(stderr, "FAIL line %d: %s\n", __LINE__, #condition); ++failures; \
} } while (0)

typedef HRESULT (WINAPI *get_class_fn)(REFCLSID, REFIID, void **);
typedef HRESULT (WINAPI *can_unload_fn)(void);

int __cdecl main(int argc, char **argv)
{
    HMODULE module;
    get_class_fn get_class;
    can_unload_fn can_unload;
    IClassFactory *factory;
    ITypeLib *typelib = NULL;
    void *object;
    WCHAR path[MAX_PATH];
    HRESULT hr;
    int i;

    if (argc != 3) {
        fprintf(stderr, "argc=%d command line=%s\n", argc, GetCommandLineA());
        fprintf(stderr, "Usage: com-smoke.exe <Excel DLL> <Excel TLB>\n");
        return 2;
    }
    setvbuf(stdout, NULL, _IONBF, 0);
    printf("LoadLibrary: %s\n", argv[1]);
    module = LoadLibraryA(argv[1]);
    if (!module) {
        fprintf(stderr, "LoadLibrary failed: %lu\n", (unsigned long)GetLastError());
        return 2;
    }
    puts("Check exports and factory");
    get_class = (get_class_fn)GetProcAddress(module, "DllGetClassObject");
    can_unload = (can_unload_fn)GetProcAddress(module, "DllCanUnloadNow");
    CHECK(get_class && can_unload);
    if (!get_class || !can_unload) goto done;
    CHECK(can_unload() == S_OK);
    object = (void *)1;
    CHECK(get_class(&CLSID_NULL, &IID_IClassFactory, &object) == CLASS_E_CLASSNOTAVAILABLE);
    CHECK(object == NULL);
    CHECK(get_class(&excel_clsid, &IID_IClassFactory, NULL) == E_POINTER);
    for (i = 0; i < 3; ++i) {
        factory = NULL;
        hr = get_class(&excel_clsid, &IID_IClassFactory, (void **)&factory);
        CHECK(hr == S_OK && factory);
        if (FAILED(hr) || !factory) break;
        CHECK(can_unload() == S_FALSE);
        object = (void *)1;
        CHECK(IClassFactory_QueryInterface(factory, &IID_IDispatch, &object) == E_NOINTERFACE);
        CHECK(object == NULL);
        object = (void *)1;
        CHECK(IClassFactory_CreateInstance(factory, (IUnknown *)factory, &IID_IDispatch, &object) == CLASS_E_NOAGGREGATION);
        CHECK(object == NULL);
        CHECK(IClassFactory_LockServer(factory, TRUE) == S_OK);
        CHECK(IClassFactory_Release(factory) == 1);
        CHECK(can_unload() == S_FALSE);
        CHECK(IClassFactory_LockServer(factory, FALSE) == S_OK);
        CHECK(can_unload() == S_OK);
    }
    CHECK(MultiByteToWideChar(CP_ACP, 0, argv[2], -1, path, MAX_PATH));
    puts("LoadTypeLibEx");
    hr = LoadTypeLibEx(path, REGKIND_NONE, &typelib);
    CHECK(hr == S_OK && typelib);
    if (typelib) {
        CHECK(ITypeLib_GetTypeInfoCount(typelib) > 0);
        ITypeLib_Release(typelib);
    }
done:
    puts("FreeLibrary");
    FreeLibrary(module);
    printf("COM smoke: %d failure(s)\n", failures);
    return failures ? 1 : 0;
}
