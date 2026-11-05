#include <windows.h>
#include <msi.h>
#include <msiquery.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>

static int failures;
#define CHECK(condition) do { if (!(condition)) { \
    fprintf(stderr, "FAIL line %d: %s\n", __LINE__, #condition); ++failures; \
} } while (0)

static MSIHANDLE query_record(MSIHANDLE database, const WCHAR *query)
{
    MSIHANDLE view = 0, record = 0;
    UINT result = MsiDatabaseOpenViewW(database, query, &view);
    if (!result) result = MsiViewExecute(view, 0);
    if (!result) result = MsiViewFetch(view, &record);
    CHECK(result == ERROR_SUCCESS);
    if (view) MsiCloseHandle(view);
    return record;
}

static void check_string(MSIHANDLE record, UINT field, const WCHAR *expected)
{
    WCHAR value[256] = {0};
    DWORD count = sizeof(value) / sizeof(value[0]);
    CHECK(MsiRecordGetStringW(record, field, value, &count) == ERROR_SUCCESS);
    CHECK(!wcscmp(value, expected));
}

int __cdecl main(int argc, char **argv)
{
    MSIHANDLE database = 0, summary = 0, record;
    WCHAR value[256] = {0};
    DWORD count;
    UINT type;
    int number = 0, bits = argc == 3 ? atoi(argv[2]) : 32;
    if ((argc != 2 && argc != 3) || (bits != 32 && bits != 64)) return 2;
    if (MsiOpenDatabaseA(argv[1], MSIDBOPEN_READONLY, &database)) return 3;
    record = query_record(database, L"SELECT `Value` FROM `Property` WHERE `Property`='Manufacturer'");
    check_string(record, 1, L"\x041e\x0410\x041e \"Etersoft\"");
    MsiCloseHandle(record);
    record = query_record(database, L"SELECT `Value` FROM `Property` WHERE `Property`='ProductCode'");
    check_string(record, 1, bits == 64 ? L"{4C7393FD-6C4A-5B3F-8EC2-B7AD16262864}" : L"{643315BA-DEF8-4CBE-AB26-C54B71535DFD}");
    MsiCloseHandle(record);
    record = query_record(database, L"SELECT `Value` FROM `Property` WHERE `Property`='UpgradeCode'");
    check_string(record, 1, bits == 64 ? L"{AB670464-4388-5009-B76D-7D4FA37EE2CF}" : L"{CD87087B-70BE-4BE2-B7F3-987AB335659D}");
    MsiCloseHandle(record);
    record = query_record(database, L"SELECT `ComponentId`, `Attributes` FROM `Component` WHERE `Component`='Ooffice'");
    check_string(record, 1, bits == 64 ? L"{F2E409F1-18B1-5FAD-ADAD-184D12608F75}" : L"{36D68DB8-D5FE-4342-B983-D189234CABC7}");
    CHECK((MsiRecordGetInteger(record, 2) & 256) == (bits == 64 ? 256 : 0));
    MsiCloseHandle(record);
    record = query_record(database, L"SELECT `Directory_Parent` FROM `Directory` WHERE `Directory`='MYDIR'");
    check_string(record, 1, bits == 64 ? L"System64Folder" : L"SystemFolder");
    MsiCloseHandle(record);
    record = query_record(database, L"SELECT `Source` FROM `CustomAction` WHERE `Action`='RegOofficeWin'");
    check_string(record, 1, bits == 64 ? L"System64Folder" : L"SystemFolder");
    MsiCloseHandle(record);
    CHECK(MsiGetSummaryInformationW(database, NULL, 0, &summary) == ERROR_SUCCESS);
    count = sizeof(value) / sizeof(value[0]);
    CHECK(MsiSummaryInfoGetPropertyW(summary, 7, &type, &number, NULL, value, &count) == ERROR_SUCCESS);
    CHECK(!wcscmp(value, bits == 64 ? L"x64;1033" : L"Intel;1033"));
    CHECK(MsiSummaryInfoGetPropertyW(summary, 14, &type, &number, NULL, NULL, NULL) == ERROR_SUCCESS);
    CHECK(number == (bits == 64 ? 200 : 100));
    count = sizeof(value) / sizeof(value[0]);
    CHECK(MsiSummaryInfoGetPropertyW(summary, 9, &type, &number, NULL, value, &count) == ERROR_SUCCESS);
    CHECK(count == 38 && value[0] == L'{' && value[37] == L'}');
    MsiCloseHandle(summary);
    MsiCloseHandle(database);
    printf("MSI%d metadata and Unicode: %d failure(s)\n", bits, failures);
    return failures ? 1 : 0;
}
