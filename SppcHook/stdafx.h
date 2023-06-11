/*
* 
* Minimum supported Windows:
*	Windows 7
* 
* Minimum supported Office:
*	Office 2013
* 
*   2013 pro plus retail: KBDNM-R8CD9-RK366-WFM3X-C7GXK
*   2016 pro plus retail: JV2QH-WNBG3-G9WFR-XRCR4-FC2QV
*   2019 pro plus retail: 73DDN-VP7FY-TQBQV-86QDJ-XW4W6
*   2021 pro plus retail: 4HV8H-CNPKG-RF622-3F7VM-FGGWK
* 
*/

#pragma once

#define DECLARE_ORIGIN(_NAME_) decltype(&(_NAME_)) Origin##_NAME_
#define DECLARE_EXTERN_ORIGIN(_NAME_) extern DECLARE_ORIGIN(_NAME_)

#include <Windows.h>
#include <slpublic.h>
#include <slerror.h>
#include <cstdio>
#include <map>

#include "Hook.h"
#include "SppcHooks.h"
#include "AdvapiHooks.h"
#include "log.h"
#include "init.h"

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

#define SPPCHOOK_NOLOG
#ifdef SPPCHOOK_NOLOG
#define printf(_format_, ...) 2
#endif
