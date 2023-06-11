#include "stdafx.h"
#pragma warning(disable:4996)

VOID InitLog(LPCSTR lpFileName) {
    assert(freopen(lpFileName, "w+", stdout));
    assert(0 == setvbuf(stdout, NULL, _IONBF, 0));
}
