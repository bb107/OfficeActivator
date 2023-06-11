#pragma once

typedef struct _SPPC_OFFICE_SKU_STATE {
	DWORD dwLicenseState;

	LPTSTR lpProductName;
	LPTSTR lpDescription;
	LPWSTR lpLicenseType;
	LPWSTR lpPartialKey;
	WCHAR SkuId[64];

}SPPC_OFFICE_SKU_STATE, * PSPPC_OFFICE_SKU_STATE;

extern HMODULE sppc;

VOID InitializeSppcFunctions();

HRESULT GetOfficeSkuState(
	_Out_ PDWORD nOfficeSkuState,
	_Outptr_result_buffer_(sizeof(SPPC_OFFICE_SKU_STATE)* (*nOfficeSkuState)) PSPPC_OFFICE_SKU_STATE* ppOfficeSkuState
);

HRESULT InstallProductKey(_In_z_ LPCWSTR lpStringProductKey);

DWORD UninstallProductKey(
	_In_ DWORD dwProductKeyCount,
	_In_ LPCWSTR* lpPartialProductKey
);

HRESULT InstallLicense(_In_z_ LPCTSTR lpLicenseFileName);

HRESULT UninstallLicense(_In_z_ LPCTSTR lpLicenseFileName);
