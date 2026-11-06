/* Verify that the published dual IID exposes HRESULT/out-parameter slots. */
#define COBJMACROS
#include <windows.h>
#include <ole2.h>
#include <stdio.h>
#include "unioffice_excel.h"

static int failures;
#define CHECK(condition) do { if (!(condition)) { \
    fprintf(stderr, "FAIL line %d: %s\n", __LINE__, #condition); ++failures; \
} } while (0)

static void creator(_Application *application)
{
    struct { LONG before; XlCreator value; LONG after; } output = {123, 0, 456};
    CHECK(_Application_get_Creator(application, &output.value) == S_OK);
    CHECK(output.value == xlCreatorCode && output.before == 123 && output.after == 456);
}

int __cdecl main(void)
{
    CLSID clsid;
    IDispatch *dispatch = NULL, *roundtrip = NULL;
    _Application *application = NULL, *created = NULL;
    IUnknown *identity = NULL, *other_identity = NULL;
    LPOLESTR quit_name = L"Quit";
    DISPID quit;
    DISPPARAMS no_arguments = {NULL, NULL, 0, 0};
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 2;
    hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    CHECK(hr == S_OK);
    if (FAILED(hr)) goto done;
    hr = CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER, &IID_IDispatch, (void**)&dispatch);
    CHECK(hr == S_OK && dispatch);
    if (!dispatch) goto done;
    CHECK(IDispatch_QueryInterface(dispatch, &IID__Application, (void**)&application) == S_OK);
    if (application) {
        creator(application);
        CHECK(_Application_QueryInterface(application, &IID_IDispatch, (void**)&roundtrip) == S_OK);
        CHECK(_Application_QueryInterface(application, &IID_IUnknown, (void**)&identity) == S_OK);
        CHECK(IDispatch_QueryInterface(dispatch, &IID_IUnknown, (void**)&other_identity) == S_OK);
        CHECK(identity && identity == other_identity);
        if (roundtrip) {
            _Application *again = NULL;
            CHECK(IDispatch_QueryInterface(roundtrip, &IID__Application, (void**)&again) == S_OK);
            CHECK(again == application);
            if (again) { creator(again); _Application_Release(again); }
        }
    }
    CHECK(CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER,
            &IID__Application, (void**)&created) == S_OK && created);
    if (created) creator(created);
done:
    if (created) _Application_Release(created);
    if (other_identity) IUnknown_Release(other_identity);
    if (identity) IUnknown_Release(identity);
    if (roundtrip) IDispatch_Release(roundtrip);
    if (application) _Application_Release(application);
    if (dispatch) {
        if (IDispatch_GetIDsOfNames(dispatch, &IID_NULL, &quit_name, 1, 0, &quit) == S_OK)
            IDispatch_Invoke(dispatch, quit, &IID_NULL, 0, DISPATCH_METHOD,
                    &no_arguments, NULL, NULL, NULL);
        IDispatch_Release(dispatch);
    }
    CoUninitialize();
    printf("Application interface: %d failure(s)\n", failures);
    return failures ? 1 : 0;
}
