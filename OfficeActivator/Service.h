#pragma once

#define SVCNAME TEXT("OfficeActivatorSvc")
#define SVCDNAME TEXT("Office Activator Service")

//
// Enable service logging
//
//#define _OA_LOG

BOOL WINAPI SvcRun();
