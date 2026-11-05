#include <windows.h>
#include <msi.h>
#include <msiquery.h>
#include <stdio.h>
#include <wchar.h>
int __cdecl main(int argc, char **argv)
{
    MSIHANDLE database = 0, view = 0, record = 0;
    WCHAR value[128];
    DWORD count = 128;
    UINT result;
    if (argc != 2) return 2;
    printf("ACP=%u\n", GetACP());
    result = MsiOpenDatabaseA(argv[1], MSIDBOPEN_READONLY, &database);
    if (result) return 3;
    result = MsiDatabaseOpenViewW(database,
        L"SELECT `Value` FROM `Property` WHERE `Property`='Manufacturer'", &view);
    if (!result) result = MsiViewExecute(view, 0);
    if (!result) result = MsiViewFetch(view, &record);
    if (!result) result = MsiRecordGetStringW(record, 1, value, &count);
    if (!result && wcscmp(value, L"\x041e\x0410\x041e \"Etersoft\"")) result = 1;
    if (record) MsiCloseHandle(record);
    if (view) MsiCloseHandle(view);
    MsiCloseHandle(database);
    printf("MSI manufacturer Unicode check: %u\n", result);
    return result ? 1 : 0;
}
