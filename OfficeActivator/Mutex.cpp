#include "pch.h"
#include "Mutex.h"
#include <sddl.h>
#pragma comment(lib,"Advapi32.lib")

#define OA_MUTEX_NAME	_T("Global\\OfficeActivatorMutex")

HANDLE WINAPI CreateGlobalMutex() {
    HANDLE mutex = nullptr;
    WCHAR* pszStringSecurityDescriptor = L"D:(A;;GA;;;WD)(A;;GA;;;AN)S:(ML;;NW;;;ME)";
    PSECURITY_DESCRIPTOR pSecDesc;

    if (ConvertStringSecurityDescriptorToSecurityDescriptor(pszStringSecurityDescriptor, SDDL_REVISION_1, &pSecDesc, NULL)) {
        SECURITY_ATTRIBUTES SecAttr;
        SecAttr.nLength = sizeof(SECURITY_ATTRIBUTES);
        SecAttr.lpSecurityDescriptor = pSecDesc;
        SecAttr.bInheritHandle = FALSE;

        mutex = CreateMutex(&SecAttr, FALSE, OA_MUTEX_NAME);
        LocalFree(pSecDesc);
    }

    return mutex;
}
