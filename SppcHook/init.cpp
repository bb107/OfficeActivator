#include "stdafx.h"

VOID Initialize() {
#ifndef SPPCHOOK_NOLOG
    InitLog("C:\\Users\\Boring\\Desktop\\console.txt");
#endif
    //InitMITM();

    InitHooks();
}
