#define STRICT

#include <windows.h>
#include <ole2.h>
#include <olenls.h>
#include <dispatch.h>

#include "test.h"

#include <pshpack1.h>
/* BSTR structure */
typedef struct
{
    ULONG clSize;
    OLECHAR abData[1];
} BYTE_BLOB16;
#include <poppack.h>
#pragma comment(lib, "ole2disp.lib");
#pragma comment(lib, "compobj.lib");
static const BYTE_BLOB16 FAR*get_blob(BSTR bstr)
{
	return ((const BYTE_BLOB16 FAR*)((char FAR*)bstr - sizeof(BYTE_BLOB16) + sizeof(OLECHAR)));
}
static void test_SysAllocStringLen(void)
{
	BSTR bstr;
	BYTE_BLOB16 *test = (BYTE_BLOB16*)"\x03\x00\x00\x00AAA";
	ok(SysStringLen(&test->abData) == 3, "SysStringLen failed\n");

	bstr = SysAllocStringLen("a", 1);
	ok(bstr != NULL, "SysAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 1, "SysAllocString failed(%ld)\n", get_blob(bstr)->clSize);
	ok(!_fmemcmp(bstr, "a", 2), "SysAllocString\n");
	ok(SysStringLen(bstr) == 1 , "SysStringLen(SysAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocStringLen("a\0", 2);
	ok(bstr != NULL, "SysAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 2, "SysAllocString failed(%ld)\n", get_blob(bstr)->clSize);
	ok(!_fmemcmp(bstr, "a\0", 3), "SysAllocString\n");
	ok(SysStringLen(bstr) == 2, "SysStringLen(SysAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocStringLen(NULL, 10);
	ok(bstr != NULL, "SysAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 10, "SysAllocStringLen failed(%ld)\n", get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 10, "SysStringLen(SysAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocStringLen("testt\0est", 9);
	ok(bstr != NULL, "SysAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 9, "SysAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(!_fmemcmp(bstr, "testt\0est", 10), "SysAllocStringLen failed\n");
	ok(SysStringLen(bstr) == 9, "SysStringLen(SysAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocStringLen("TESTTEST", 3);
	ok(bstr != NULL, "SysAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 3, "SysAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(!_fmemcmp(bstr, "TES", 4), "SysAllocStringLen failed\n");
	ok(SysStringLen(bstr) == 3, "SysStringLen(SysAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);
}
static void test_SysAllocString(void)
{
	BSTR bstr;
	bstr = SysAllocString("testtest");
	ok(bstr != NULL, "SysAllocString failed\n");
	ok(get_blob(bstr)->clSize == 8, "SysAllocString failed(%ld)\n", get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 8, "SysStringLen(SysAllocString) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("testt\0est");
	ok(bstr != NULL, "SysAllocString failed\n");
	ok(get_blob(bstr)->clSize == 5, "SysAllocString failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(!_fmemcmp(bstr, "testt\0est", 5), "SysAllocString\n");
	ok(SysStringLen(bstr) == 5, "SysStringLen(SysAllocString) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);
}

static void test_SysReAllocString(void)
{
	BSTR bstr;
	bstr = SysAllocString("test");
	SysReAllocString((BSTR FAR*)&bstr, "TESTA");
	ok(!_fmemcmp(bstr, "TESTA", 6), "SysReAllocStirng failed\n");
	ok(get_blob(bstr)->clSize == 5, "SysReAllocStirng failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 5, "SysStringLen(SysReAllocStirng) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocString((BSTR FAR*)&bstr, NULL);
	ok(bstr == NULL, "SysReAllocStirng(&bstr, NULL) must be fail\n");
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocString((BSTR FAR*)&bstr, "A");
	ok(!_fmemcmp(bstr, "A", 2), "SysReAllocStirng failed\n");
	ok(get_blob(bstr)->clSize == 1, "SysReAllocStirng failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 1, "SysStringLen(SysReAllocStirng) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocString((BSTR FAR*)&bstr, "A\0B");
	ok(!_fmemcmp(bstr, "A\0B", 2), "SysReAllocStirng failed\n");
	ok(get_blob(bstr)->clSize == 1, "SysReAllocStirng failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 1, "SysStringLen(SysReAllocStirng) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);
}
static void test_SysReAllocStringLen(void)
{
	BSTR bstr;
	bstr = SysAllocString("test");
	SysReAllocStringLen((BSTR FAR*)&bstr, "TESTA", 5);
	ok(!_fmemcmp(bstr, "TESTA", 6), "SysReAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 5, "SysReAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 5, "SysStringLen(SysReAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocStringLen((BSTR FAR*)&bstr, NULL, 1);
	ok(bstr != NULL, "SysReAllocStringLen(&bstr, NULL, 1) failed\n");
	ok(get_blob(bstr)->clSize == 1, "SysReAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 1, "SysStringLen(SysReAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocStringLen((BSTR FAR*)&bstr, "A", 1);
	ok(!_fmemcmp(bstr, "A", 2), "SysReAllocStirng failed\n");
	ok(get_blob(bstr)->clSize == 1, "SysReAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 1, "SysStringLen(SysReAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocStringLen((BSTR FAR*)&bstr, "A\0B", 3);
	ok(!_fmemcmp(bstr, "A\0B", 4), "SysReAllocStirng failed\n");
	ok(get_blob(bstr)->clSize == 3, "SysReAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(SysStringLen(bstr) == 3, "SysStringLen(SysReAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);

	bstr = SysAllocString("test");
	SysReAllocStringLen((BSTR FAR*)&bstr, "TESTTEST", 3);
	ok(bstr != NULL, "SysReAllocStringLen failed\n");
	ok(get_blob(bstr)->clSize == 3, "SysReAllocStringLen failed(%ld)\n", (DWORD)get_blob(bstr)->clSize);
	ok(!_fmemcmp(bstr, "TES", 4), "SysReAllocStringLen failed\n");
	ok(SysStringLen(bstr) == 3, "SysReAllocStringLen(SysAllocStringLen) failed(%ld)\n", (DWORD)SysStringLen(bstr));
	SysFreeString(bstr);
}
START_TEST(ole2disp)
{
	CoInitialize(NULL);
	test_SysAllocStringLen();
	test_SysAllocString();
	test_SysReAllocString();
	test_SysReAllocStringLen();
}
