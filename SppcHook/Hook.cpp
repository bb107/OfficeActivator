#include "stdafx.h"
#include "Detours/detours.h"
#include <tchar.h>
#pragma comment(lib, "ntdll.lib")

HMODULE sppc;
HMODULE advapi;

#define BIND_HOOK(_NAME_, _LIB_) BindHook((PVOID*)&(Origin##_NAME_), #_NAME_, Hook##_NAME_, _LIB_)
#define SPPC_BIND_HOOK(_NAME_) BIND_HOOK(_NAME_, sppc)
#define CRYPT_BIND_HOOK(_NAME_) BIND_HOOK(_NAME_, advapi)

extern "C" {
	NTSYSAPI VOID NTAPI RtlGetNtVersionNumbers(
		_Out_opt_ DWORD* MajorVersion,
		_Out_opt_ DWORD* MinorVersion,
		_Out_opt_ DWORD* BuildNumber
	);
}

void BindHook(PVOID* Origin, PCSTR Name, PVOID Dest, HMODULE hm) {
    assert(*Origin = GetProcAddress(hm, Name));
    assert(0 == DetourAttach(Origin, Dest));
}

BOOL RegReadValue(
	_Out_ PDWORD pcbData,
	_Out_ PBYTE* ppbData,
	_In_ HKEY hKey,
	_In_ LPCTSTR pszSubKey,
	_In_opt_ LPCTSTR pszValueName) {
	HKEY hSubKey;
	BOOL success = FALSE;
	PBYTE pData = nullptr;

	*pcbData = 0;
	*ppbData = nullptr;

	DWORD dwKeyType;
	DWORD dwDataSize = 0;
	LSTATUS result;

	result = RegOpenKey(
		hKey,
		pszSubKey,
		&hSubKey
	);
	if (ERROR_SUCCESS == result) {
		result = RegQueryValueEx(hSubKey,
			pszValueName,
			nullptr,
			&dwKeyType,
			pData,
			&dwDataSize
		);
		if (result == ERROR_FILE_NOT_FOUND ||
			dwKeyType != REG_SZ ||
			dwDataSize == 0) {
			//
			// we failed.
			//
		}
		else {
			pData = (PBYTE)LocalAlloc(0, dwDataSize);
			assert(pData);

			result = RegQueryValueEx(hSubKey,
				pszValueName,
				nullptr,
				&dwKeyType,
				pData,
				&dwDataSize
			);

			success = result == ERROR_SUCCESS;
		}

		RegCloseKey(hSubKey);
	}

	if (success) {
		*pcbData = dwDataSize;
		*ppbData = pData;
	}
	else {
		LocalFree(pData);
	}

	return success;
}

VOID InitHooks() {
    assert(0 == DetourTransactionBegin());
    assert(0 == DetourUpdateThread(GetCurrentThread()));

	DWORD dwValue;
	LPTSTR lpValue;

	DWORD MajorVersion, MinorVersion;
	RtlGetNtVersionNumbers(&MajorVersion, &MinorVersion, nullptr);

	if (MajorVersion == 6 && MinorVersion == 1) {
		if (RegReadValue(
			&dwValue,
			(PBYTE*)&lpValue,
			HKEY_LOCAL_MACHINE,
			_T("SOFTWARE\\Microsoft\\OfficeSoftwareProtectionPlatform"),
			_T("Path"))) {
			dwValue += sizeof(TCHAR) * 9;

			LPTSTR pOsppc = (LPTSTR)LocalAlloc(0, dwValue);
			assert(pOsppc);

			_tcscpy_s(pOsppc, dwValue / sizeof(TCHAR), lpValue);
			_tcscat_s(pOsppc, dwValue / sizeof(TCHAR), _T("osppc.dll"));

			sppc = LoadLibrary(pOsppc);
			LocalFree(pOsppc);
			LocalFree(lpValue);
		}
		else {
			//
			// we failed.
			//
		}
	}
	else {
		sppc = LoadLibraryA("sppc.dll");
	}

    assert(sppc);
    SPPC_BIND_HOOK(SLGetLicensingStatusInformation);
    //SPPC_BIND_HOOK(SLGetApplicationPolicy);
    //SPPC_BIND_HOOK(SLGetSLIDList);
    //SPPC_BIND_HOOK(SLGetPolicyInformation);
    //SPPC_BIND_HOOK(SLSetAuthenticationData);
    //SPPC_BIND_HOOK(SLGetAuthenticationResult);
    //SPPC_BIND_HOOK(SLGetProductSkuInformation);
    
    //advapi = LoadLibraryA("Advapi32.dll");
    //assert(advapi);
    //CRYPT_BIND_HOOK(CryptImportKey);

    assert(0 == DetourTransactionCommit());
}
