#include "pch.h"
#include "Sppc.h"
#include "Helps.h"
#include <slpublic.h>
#include <slerror.h>
#pragma comment(lib,"Rpcrt4.lib")

decltype(&SLLoadApplicationPolicies)__SLLoadApplicationPolicies;
decltype(&SLGetApplicationPolicy)__SLGetApplicationPolicy;
decltype(&SLUnloadApplicationPolicies)__SLUnloadApplicationPolicies;
decltype(&SLOpen)__SLOpen;
decltype(&SLGetSLIDList)__SLGetSLIDList;
decltype(&SLGetLicensingStatusInformation)__SLGetLicensingStatusInformation;
decltype(&SLGetProductSkuInformation)__SLGetProductSkuInformation;
decltype(&SLGetPKeyInformation)__SLGetPKeyInformation;
decltype(&SLClose)__SLClose;
decltype(&SLInstallProofOfPurchase)__SLInstallProofOfPurchase;
decltype(&SLUninstallProofOfPurchase)__SLUninstallProofOfPurchase;
decltype(&SLInstallLicense)__SLInstallLicense;
decltype(&SLGetLicenseFileId)__SLGetLicenseFileId;
decltype(&SLUninstallLicense)__SLUninstallLicense;

SLID officeAppId;
HMODULE sppc = nullptr;

extern "C" {
	NTSYSAPI VOID NTAPI RtlGetNtVersionNumbers(
		_Out_opt_ DWORD* MajorVersion,
		_Out_opt_ DWORD* MinorVersion,
		_Out_opt_ DWORD* BuildNumber
	);
}

void GuidToString(
	_In_ GUID* guid,
	_In_ DWORD cbString,
	_Out_ LPWSTR lpString) {
	swprintf_s(
		lpString, cbString,
		L"{%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
		guid->Data1, guid->Data2, guid->Data3,
		guid->Data4[0], guid->Data4[1], guid->Data4[2], guid->Data4[3],
		guid->Data4[4], guid->Data4[5], guid->Data4[6], guid->Data4[7]
	);
}

VOID InitializeSppcFunctions() {
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
			ASSERT(pOsppc);

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

	ASSERT(sppc);

	__SLLoadApplicationPolicies = (decltype(&SLLoadApplicationPolicies))GetProcAddress(sppc, "SLLoadApplicationPolicies");
	__SLGetApplicationPolicy = (decltype(&SLGetApplicationPolicy))GetProcAddress(sppc, "SLGetApplicationPolicy");
	__SLUnloadApplicationPolicies = (decltype(&SLUnloadApplicationPolicies))GetProcAddress(sppc, "SLUnloadApplicationPolicies");
	__SLOpen = (decltype(&SLOpen))GetProcAddress(sppc, "SLOpen");
	__SLGetSLIDList = (decltype(&SLGetSLIDList))GetProcAddress(sppc, "SLGetSLIDList");
	__SLGetLicensingStatusInformation = (decltype(&SLGetLicensingStatusInformation))GetProcAddress(sppc, "SLGetLicensingStatusInformation");
	__SLGetProductSkuInformation = (decltype(&SLGetProductSkuInformation))GetProcAddress(sppc, "SLGetProductSkuInformation");
	__SLGetPKeyInformation = (decltype(&SLGetPKeyInformation))GetProcAddress(sppc, "SLGetPKeyInformation");
	__SLClose = (decltype(&SLClose))GetProcAddress(sppc, "SLClose");
	__SLInstallProofOfPurchase = (decltype(&SLInstallProofOfPurchase))GetProcAddress(sppc, "SLInstallProofOfPurchase");
	__SLUninstallProofOfPurchase = (decltype(&SLUninstallProofOfPurchase))GetProcAddress(sppc, "SLUninstallProofOfPurchase");
	__SLInstallLicense = (decltype(&SLInstallLicense))GetProcAddress(sppc, "SLInstallLicense");
	__SLGetLicenseFileId = (decltype(&SLGetLicenseFileId))GetProcAddress(sppc, "SLGetLicenseFileId");
	__SLUninstallLicense = (decltype(&SLUninstallLicense))GetProcAddress(sppc, "SLUninstallLicense");

	//BOOL success = SUCCEEDED(CLSIDFromString(L"{59a52881-a989-479d-af46-f275c6370663}", &officeAppId));//Office 2010
	BOOL success = SUCCEEDED(CLSIDFromString(L"{0ff1ce15-a989-479d-af46-f275c6370663}", &officeAppId));
	ASSERT(success);
}

HRESULT GetOfficeSkuState(
	_Out_ PDWORD nOfficeSkuState,
	_Outptr_result_buffer_(sizeof(SPPC_OFFICE_SKU_STATE)* (*nOfficeSkuState)) PSPPC_OFFICE_SKU_STATE* ppOfficeSkuState) {
	HSLC hslc;
	UINT nSkus = 0;
	PSPPC_OFFICE_SKU_STATE result = nullptr;
	HRESULT hr = ERROR_SUCCESS;

	*nOfficeSkuState = 0;
	*ppOfficeSkuState = nullptr;

	hr = __SLOpen(&hslc);
	if (SUCCEEDED(hr)) {
		SLID* pSkus;

		hr = __SLGetSLIDList(hslc, SL_ID_APPLICATION, &officeAppId, SL_ID_PRODUCT_SKU, &nSkus, &pSkus);
		if (SUCCEEDED(hr)) {

			result = (PSPPC_OFFICE_SKU_STATE)LocalAlloc(LMEM_ZEROINIT, sizeof(SPPC_OFFICE_SKU_STATE) * nSkus);
			ASSERT(result);

			for (DWORD i = 0; i < nSkus; ++i) {
				HSLP hPolicy;
				SPPC_OFFICE_SKU_STATE& state = result[i];

				if (SUCCEEDED(__SLLoadApplicationPolicies(&officeAppId, &pSkus[i], 0, &hPolicy))) {
					UINT cbData;
					PBYTE pbData;
					if (SUCCEEDED(__SLGetApplicationPolicy(hPolicy, L"office-LicenseType", nullptr, &cbData, &pbData))) {
						state.lpLicenseType = (LPWSTR)pbData;
					}

					__SLUnloadApplicationPolicies(hPolicy, 0);
				}

				UINT nLicensingStatus;
				SL_LICENSING_STATUS* ppLicensingStatus;
				if (SUCCEEDED(__SLGetLicensingStatusInformation(hslc, &officeAppId, &pSkus[i], nullptr, &nLicensingStatus, &ppLicensingStatus))) {
					ASSERT(nLicensingStatus == 1);

					state.dwLicenseState = (DWORD)ppLicensingStatus[0].eStatus;
					LocalFree(ppLicensingStatus);
				}

				UINT nName, nDescription;
				PBYTE pName, pDescription;
				if (SUCCEEDED(__SLGetProductSkuInformation(hslc, &pSkus[i], L"productName", nullptr, &nName, &pName))) {
					state.lpProductName = (LPTSTR)pName;
				}

				if (SUCCEEDED(__SLGetProductSkuInformation(hslc, &pSkus[i], L"productDescription", nullptr, &nDescription, &pDescription))) {
					state.lpDescription = (LPTSTR)pDescription;
				}

				UINT nKeys;
				SLID* pKeys;
				if (SUCCEEDED(__SLGetSLIDList(hslc, SL_ID_PRODUCT_SKU, &pSkus[i], SL_ID_PKEY, &nKeys, &pKeys))) {

					ASSERT(nKeys == 1);

					UINT cbKey;
					PBYTE pbKey;
					if (SUCCEEDED(__SLGetPKeyInformation(hslc, pKeys, SL_INFO_KEY_PARTIAL_PRODUCT_KEY, nullptr, &cbKey, &pbKey))) {
						state.lpPartialKey = (LPWSTR)pbKey;
					}

					LocalFree(pKeys);
				}

				GuidToString(&pSkus[i], sizeof(state.SkuId) / sizeof(WCHAR), state.SkuId);
			}

			LocalFree(pSkus);
		}

		__SLClose(hslc);
	}

	*nOfficeSkuState = nSkus;
	*ppOfficeSkuState = result;
	return hr;
}

HRESULT InstallProductKey(_In_z_ LPCWSTR lpStringProductKey) {
	SLID keyId;
	HSLC hSlc;
	HRESULT hr;

	hr = __SLOpen(&hSlc);
	if (SUCCEEDED(hr)) {
		hr = __SLInstallProofOfPurchase(
			hSlc,
			SL_PKEY_DETECT,
			lpStringProductKey,
			0,
			nullptr,
			&keyId
		);

		__SLClose(hSlc);
	}

	return hr;
}

DWORD UninstallProductKey(
	_In_ DWORD dwProductKeyCount,
	_In_ LPCWSTR* lpPartialProductKey) {
	DWORD result = 0;
	HSLC hslc;
	HRESULT hr = __SLOpen(&hslc);
	if (SUCCEEDED(hr)) {

		UINT nSku;
		SLID* pSku;
		if (SUCCEEDED(__SLGetSLIDList(hslc, SL_ID_APPLICATION, &officeAppId, SL_ID_PRODUCT_SKU, &nSku, &pSku))) {

			for (UINT i = 0; i < nSku; ++i) {
				UINT nPkey;
				SLID* pPkey;
				if (SUCCEEDED(__SLGetSLIDList(hslc, SL_ID_PRODUCT_SKU, &pSku[i], SL_ID_PKEY, &nPkey, &pPkey))) {

					for (UINT j = 0; j < nPkey; ++j) {
						UINT cbPartialKey;
						LPWSTR pbPartialKey;
						if (SUCCEEDED(__SLGetPKeyInformation(hslc, &pPkey[j], SL_INFO_KEY_PARTIAL_PRODUCT_KEY, nullptr, &cbPartialKey, (PBYTE*)&pbPartialKey))) {

							for (DWORD k = 0; k < dwProductKeyCount; ++k) {
								if (_wcsnicmp(lpPartialProductKey[k], pbPartialKey, 5) == 0) {
									if (SUCCEEDED(__SLUninstallProofOfPurchase(hslc, &pPkey[j]))) {
										++result;
										break;
									}
								}
							}

							LocalFree(pbPartialKey);
						}
					}

					LocalFree(pPkey);
				}
			}

			LocalFree(pSku);
		}

		__SLClose(hslc);
	}

	return result;
}

HRESULT InstallLicense(_In_z_ LPCTSTR lpLicenseFileName) {

	HRESULT hr;
	HSLC hslc = nullptr;
	SLID id;

	DWORD dwFileSize;
	PBYTE pbData = (PBYTE)MapFileReadOnly(lpLicenseFileName, &dwFileSize);
	if (!pbData)return SL_E_NOT_SUPPORTED;

	do {
		hr = __SLOpen(&hslc);
		if (!SUCCEEDED(hr))break;

		hr = __SLInstallLicense(hslc, dwFileSize, pbData, &id);

		__SLClose(hslc);
	} while (FALSE);

	UnmapViewOfFile(pbData);
	return hr;
}

HRESULT UninstallLicense(_In_z_ LPCTSTR lpLicenseFileName) {

	HRESULT hr;
	HSLC hslc = nullptr;
	SLID id;

	DWORD dwFileSize;
	PBYTE pbData = (PBYTE)MapFileReadOnly(lpLicenseFileName, &dwFileSize);
	if (!pbData)return SL_E_NOT_SUPPORTED;

	do {
		hr = __SLOpen(&hslc);
		if (!SUCCEEDED(hr))break;

		hr = __SLGetLicenseFileId(hslc, dwFileSize, pbData, &id);
		if (!SUCCEEDED(hr))break;

		hr = __SLUninstallLicense(hslc, &id);

	} while (FALSE);

	if (hslc)__SLClose(hslc);
	UnmapViewOfFile(pbData);
	return hr;
}
