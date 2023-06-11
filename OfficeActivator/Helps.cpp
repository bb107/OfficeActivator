#include "pch.h"
#include "Helps.h"
#include "Sppc.h"
#include <WinTrust.h>
#include <Softpub.h>
#pragma comment(lib,"wintrust.lib")

BOOL RegReadValue(
	_Out_ PDWORD pcbData,
	_Out_ PBYTE* ppbData,
	_In_ HKEY hKey,
	_In_ LPCTSTR pszSubKey,
	_In_opt_ LPCTSTR pszValueName,
	_In_ DWORD dwSamDesired) {
	HKEY hSubKey;
	BOOL success = FALSE;
	PBYTE pData = nullptr;

	*pcbData = 0;
	*ppbData = nullptr;

	DWORD dwKeyType;
	DWORD dwDataSize = 0;
	LSTATUS result;

	result = RegOpenKeyEx(
		hKey,
		pszSubKey,
		0,
		dwSamDesired | KEY_READ,
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
			ASSERT(pData);

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

BOOL GetOfficeInstallationPath(_Out_ LPTSTR* lpOfficePath) {
	DWORD dwSize;
	LPTSTR _lpOfficePath;

	*lpOfficePath = nullptr;

	BOOL result;
	result = RegReadValue(
		&dwSize,
		(PBYTE*)&_lpOfficePath,
		HKEY_LOCAL_MACHINE,
		_T("SOFTWARE\\Microsoft\\Office\\15.0\\Common\\InstallRoot"),
		_T("Path")
	);
	if (result) {
		*lpOfficePath = _lpOfficePath;
		return result;
	}

	result = RegReadValue(
		&dwSize,
		(PBYTE*)&_lpOfficePath,
		HKEY_LOCAL_MACHINE,
		_T("SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRunStore\\Applications"),
		nullptr
	);
	if (result) {
		*lpOfficePath = _lpOfficePath;
		return result;
	}

	result = RegReadValue(
		&dwSize,
		(PBYTE*)&_lpOfficePath,
		HKEY_LOCAL_MACHINE,
		_T("SOFTWARE\\Microsoft\\Office\\16.0\\ClickToRunStore\\Applications"),
		nullptr
	);
	if (result) {
		*lpOfficePath = _lpOfficePath;
		return result;
	}

	return result;
}

BOOL GetMsoPath(_Out_ LPTSTR* lpMsoPath) {
	DWORD dwSize;
	LPTSTR _lpMsoPath;

	*lpMsoPath = nullptr;

	BOOL result;
	result = RegReadValue(
		&dwSize,
		(PBYTE*)&_lpMsoPath,
		HKEY_LOCAL_MACHINE,
		_T("SOFTWARE\\Microsoft\\Office\\15.0\\Common\\FilesPaths"),
		_T("mso.dll")
	);
	if (result) {
		*lpMsoPath = _lpMsoPath;
		return result;
	}

	result = RegReadValue(
		&dwSize,
		(PBYTE*)&_lpMsoPath,
		HKEY_LOCAL_MACHINE,
		_T("SOFTWARE\\Microsoft\\Office\\15.0\\Common\\FilesPaths"),
		_T("office.odf")
	);
	if (result) {
		result = FALSE;

		BOOL last = FALSE;
		DWORD count = 0;
		LPTSTR end = _lpMsoPath + (dwSize / sizeof(TCHAR)) - 2;
		while (end > _lpMsoPath) {
			if (*end == _T('\\')) {

				if (last == FALSE) {
					last = TRUE;
				}
				else {
					ASSERT(count >= 7);
					_tcscpy_s(end + 1, count, _T("MSO.DLL"));

					result = TRUE;
					break;
				}
			}

			++count;
			--end;
		}

		if (result) {
			*lpMsoPath = _lpMsoPath;
		}
		else {
			LocalFree(_lpMsoPath);
		}

		return result;
	}

	result = RegReadValue(
		&dwSize,
		(PBYTE*)&_lpMsoPath,
		HKEY_LOCAL_MACHINE,
		_T("SOFTWARE\\Microsoft\\Office\\16.0\\Common"),
		_T("MsoWerCrashDllPath")
	);
	if (result) {

		result = FALSE;

		DWORD count = 0;
		LPTSTR end = _lpMsoPath + (dwSize / sizeof(TCHAR)) - 2;
		while (end > _lpMsoPath) {
			if (*end == _T('\\')) {

				ASSERT(count >= 7);
				_tcscpy_s(end + 1, count, _T("MSO.DLL"));

				result = TRUE;
				break;
			}

			++count;
			--end;
		}

		if (result) {
			*lpMsoPath = _lpMsoPath;
		}
		else {
			LocalFree(_lpMsoPath);
		}

		return result;
	}

	return result;
}

BOOL VerifyEmbeddedSignature(LPCWSTR pwszSourceFile)
{
	LONG lStatus;
	BOOL result;

	WINTRUST_FILE_INFO FileData;
	memset(&FileData, 0, sizeof(FileData));
	FileData.cbStruct = sizeof(WINTRUST_FILE_INFO);
	FileData.pcwszFilePath = pwszSourceFile;
	FileData.hFile = NULL;
	FileData.pgKnownSubject = NULL;

	GUID WVTPolicyGUID = WINTRUST_ACTION_GENERIC_VERIFY_V2;
	WINTRUST_DATA WinTrustData;

	memset(&WinTrustData, 0, sizeof(WinTrustData));

	WinTrustData.cbStruct = sizeof(WinTrustData);
	WinTrustData.pPolicyCallbackData = NULL;
	WinTrustData.pSIPClientData = NULL;
	WinTrustData.dwUIChoice = WTD_UI_NONE;
	WinTrustData.fdwRevocationChecks = WTD_REVOKE_NONE;
	WinTrustData.dwUnionChoice = WTD_CHOICE_FILE;
	WinTrustData.dwStateAction = WTD_STATEACTION_VERIFY;
	WinTrustData.hWVTStateData = NULL;
	WinTrustData.pwszURLReference = NULL;
	WinTrustData.dwUIContext = 0;
	WinTrustData.pFile = &FileData;

	result = ERROR_SUCCESS == WinVerifyTrust(
		NULL,
		&WVTPolicyGUID,
		&WinTrustData
	);

	WinTrustData.dwStateAction = WTD_STATEACTION_CLOSE;

	lStatus = WinVerifyTrust(
		NULL,
		&WVTPolicyGUID,
		&WinTrustData);

	return result;
}

CString SLErrorToString(HRESULT error) {
	LPTSTR errorText;
	FormatMessage(
		FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
		sppc,
		error,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPTSTR)&errorText,
		0,
		NULL
	);

	CString msg = errorText;
	LocalFree(errorText);

	return msg;
}

PVOID MapFileReadOnly(
	_In_ LPCTSTR lpFilePath,
	_Out_opt_ PDWORD lpFileSize) {

	PVOID result = nullptr;

	if (lpFileSize)*lpFileSize = 0;

	HANDLE hFile = CreateFile(
		lpFilePath,
		GENERIC_READ,
		FILE_SHARE_READ,
		nullptr,
		OPEN_EXISTING,
		0,
		nullptr
	);
	if (hFile != INVALID_HANDLE_VALUE) {

		if (lpFileSize) {
			*lpFileSize = GetFileSize(hFile, nullptr);
		}

		HANDLE hMapping = CreateFileMapping(
			hFile,
			nullptr,
			PAGE_READONLY | SEC_COMMIT,
			0,
			0,
			nullptr
		);
		if (hMapping) {
			result = MapViewOfFile(
				hMapping,
				FILE_MAP_READ,
				0,
				0,
				0
			);

			CloseHandle(hMapping);
		}


		CloseHandle(hFile);
	}

	return result;
}
