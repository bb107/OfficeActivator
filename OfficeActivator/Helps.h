#pragma once

BOOL RegReadValue(
	_Out_ PDWORD pcbData,
	_Out_ PBYTE* ppbData,
	_In_ HKEY hKey,
	_In_ LPCTSTR pszSubKey,
	_In_opt_ LPCTSTR pszValueName,
	_In_ DWORD dwSamDesired = 0
);

PVOID MapFileReadOnly(
	_In_ LPCTSTR lpFilePath,
	_Out_opt_ PDWORD lpFileSize = nullptr
);

BOOL GetOfficeInstallationPath(_Out_ LPTSTR* lpOfficePath);

BOOL GetMsoPath(_Out_ LPTSTR* lpMsoPath);

BOOL VerifyEmbeddedSignature(LPCWSTR pwszSourceFile);

CString SLErrorToString(HRESULT error);
