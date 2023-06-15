#include "pch.h"
#include "framework.h"
#include "OfficeActivator.h"
#include "OfficeActivatorDlg.h"
#include "afxdialogex.h"
#include "Helps.h"
#include "PeFile.h"
#include "Sppc.h"
#include "Mutex.h"
#include <vector>
#include <map>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define SPPCHOOK_NOT_FOUND 1
#define MSO_NOT_PATCHED 2
#define INVALID_MSO_STATE 4

#define OSP_LICENSE_PPD 0x00000001
#define OSP_LICENSE_UL_OOB 0x00000002
#define OSP_LICENSE_UL_PHN 0x00000004
#define OSP_LICENSE_UL 0x00000008
#define OSP_LICENSE_PL 0x00000010

#ifdef _WIN64
#define SPPC_NAME _T("SppcHook64.dll")
#else
#define SPPC_NAME _T("SppcHook32.dll")
#endif

LPCTSTR g_TypeTable[] = {
	_T("ppd.xrm-ms"),
	_T("ul-oob.xrm-ms"),
	_T("ul-phn.xrm-ms"),
	_T("ul.xrm-ms"),
	_T("pl.xrm-ms"),
};
HANDLE g_Mutex;

typedef struct _OSP_LICENSE {

	DWORD dwFlags;
	DWORD dwReserved;

	LPTSTR lpLicenseName;

}OSP_LICENSE, * POSP_LICENSE;

struct cmp_str {
	bool operator()(LPCTSTR a, LPCTSTR b) const {
		return _tcsicmp(a, b) < 0;
	}
};

std::vector<std::pair<WORD, std::map<LPCTSTR, OSP_LICENSE, cmp_str>>> g_LicenseList;

VOID COfficeActivatorDlg::AutoDetectOffice()
{
	LPTSTR lpOfficePath;
	LPTSTR lpMsoPath;

	GetOfficeInstallationPath(&lpOfficePath);
	GetMsoPath(&lpMsoPath);

	if (lpMsoPath && lpOfficePath) {
		SetDlgItemText(IDC_EDIT_OFFICE_PATH, lpOfficePath);
		SetDlgItemText(IDC_EDIT_MSO_PATH, lpMsoPath);
	}

	LocalFree(lpOfficePath);
	LocalFree(lpMsoPath);
}

VOID InitializeListControl(CListCtrl* list) {
	//
	// Show grid lines
	//
	list->SetExtendedStyle(list->GetExtendedStyle() | LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES);

	int nCol = 0;

	list->InsertColumn(nCol++, _T("ProductName"), LVCFMT_LEFT, 800);
	list->InsertColumn(nCol++, _T("Description"), LVCFMT_LEFT, 500);
	list->InsertColumn(nCol++, _T("LicenseType"), LVCFMT_LEFT, 200);
	//list->InsertColumn(nCol++, _T("LicenseState"), LVCFMT_LEFT, 200);
	list->InsertColumn(nCol++, _T("PartialKey"), LVCFMT_LEFT, 200);
	list->InsertColumn(nCol++, _T("SkuId"), LVCFMT_LEFT, 600);
}

VOID InsertListItem(CListCtrl* list, PSPPC_OFFICE_SKU_STATE state) {
	INT nItem = list->InsertItem(list->GetItemCount(), state->lpProductName);
	int nCol = 1;

	list->SetItemText(nItem, nCol++, state->lpDescription);
	list->SetItemText(nItem, nCol++, state->lpLicenseType);

	//CString s;
	//s.Format(_T("%d"), state->dwLicenseState);
	//list->SetItemText(nItem, nCol++, s);

	list->SetItemText(nItem, nCol++, state->lpPartialKey);
	list->SetItemText(nItem, nCol++, state->SkuId);
}

VOID UpdateLicenseList(CListCtrl* list) {
	list->DeleteAllItems();

	DWORD size;
	PSPPC_OFFICE_SKU_STATE state;
	if (SUCCEEDED(GetOfficeSkuState(&size, &state))) {

		for (DWORD i = 0; i < size; ++i) {
			InsertListItem(list, &state[i]);

			LocalFree(state[i].lpProductName);
			LocalFree(state[i].lpDescription);
			LocalFree(state[i].lpLicenseType);
			LocalFree(state[i].lpPartialKey);
		}

		LocalFree(state);
	}

	for (int i = 0; i < list->GetHeaderCtrl()->GetItemCount(); ++i)
		list->SetColumnWidth(i, LVSCW_AUTOSIZE_USEHEADER);
}

VOID InitializeProductKeyList(CComboBox* list) {
	list->SetItemDataPtr(list->AddString(_T("Office 2013 ProPlus MSDN Retail")), _T("KBDNM-R8CD9-RK366-WFM3X-C7GXK"));
	list->SetItemDataPtr(list->AddString(_T("Office 2016 ProPlus Retail")), _T("JV2QH-WNBG3-G9WFR-XRCR4-FC2QV"));
	list->SetItemDataPtr(list->AddString(_T("Office 2016 Professional Retail")), _T("D77NY-XJ4WK-3MGPX-FGBYF-QGPX7"));
	list->SetItemDataPtr(list->AddString(_T("Office 2019 ProPlus MSDN Retail")), _T("73DDN-VP7FY-TQBQV-86QDJ-XW4W6"));
	list->SetItemDataPtr(list->AddString(_T("Office 2021 ProPlus MSDN Retail")), _T("4HV8H-CNPKG-RF622-3F7VM-FGGWK"));
}

VOID InitializeLicenseDirectoryMap() {
	WIN32_FIND_DATA wfd{};
	HANDLE hff = FindFirstFile(_T("licenses\\*"), &wfd);

	if (hff == INVALID_HANDLE_VALUE) {
		MessageBox(nullptr, _T("Can not found licenses folder."), _T("Warning"), MB_OK);
		return;
	}
	else {

		do {

			if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY && wfd.cFileName[0] != _T('.')) {

				ULONG value = _tcstoul(wfd.cFileName, nullptr, 10);
				WORD wNumber = value & 0xffff;
				if ((value & ~0xffff) || wNumber > 2021 || wNumber < 2013) {
					//
					// Invalid version
					//
				}
				else {
					std::map<LPCTSTR, OSP_LICENSE, cmp_str> map;
					TCHAR buffer[10 + 4 + 9] = _T("licenses\\");

					_tcscat_s(buffer, wfd.cFileName);
					_tcscat_s(buffer, _T("\\*.xrm-ms"));

					WIN32_FIND_DATA wfd2{};
					HANDLE hff2 = FindFirstFile(buffer, &wfd2);
					if (hff2 == INVALID_HANDLE_VALUE) {
						//
						// Empty folder?
						//
					}
					else {

						do {

							CString fileName = wfd2.cFileName;
							int index = fileName.Find(_T('-'));
							if (index == -1) {
								//
								// Invalid file name format
								//
							}
							else {
								CString name = fileName.Left(index);
								CString type = fileName.Mid(index + 1);
								DWORD dwFlags = 0;

								for (DWORD i = 0; i < sizeof(g_TypeTable) / sizeof(LPCTSTR); ++i) {
									if (type == g_TypeTable[i]) {
										dwFlags = 1 << i;
										break;
									}
								}

								if (dwFlags == 0) {
									//
									// Invalid file type
									//
									continue;
								}

								auto iter = map.find(name);
								if (iter == map.end()) {
									DWORD length = name.GetLength() + 1;
									OSP_LICENSE license;

									license.dwReserved = 0;
									license.dwFlags = dwFlags;
									license.lpLicenseName = (LPTSTR)LocalAlloc(0, sizeof(TCHAR) * length);
									ASSERT(license.lpLicenseName);

									_tcscpy_s(license.lpLicenseName, length, name);

									map.insert(std::make_pair(license.lpLicenseName, license));
								}
								else {
									if (iter->second.dwFlags & dwFlags) {
										//
										// fatal error
										//

										ASSERT(FALSE);
									}
									else {
										iter->second.dwFlags |= dwFlags;
									}
								}
							}

						} while (FindNextFile(hff2, &wfd2));

						FindClose(hff2);
					}

					if (map.empty()) {
						//
						// No valid license file found.
						//
					}
					else {
						g_LicenseList.push_back(std::make_pair(wNumber, map));
					}
				}

			}

		} while (FindNextFile(hff, &wfd));

		FindClose(hff);
	}

}

VOID InitializeLicenseFileControls(CComboBox* cbx) {

	cbx->ResetContent();

	for (auto iter = g_LicenseList.begin(); iter != g_LicenseList.end(); ++iter) {
		CString version;
		version.Format(_T("%d"), iter->first);

		cbx->AddString(version);
	}

}

COfficeActivatorDlg::COfficeActivatorDlg(CWnd* pParent)
	: CDialogEx(IDD_OFFICEACTIVATOR_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_PatchState = ~0;

	m_RedColor = RGB(255, 0, 0);
	m_GreenColor = RGB(0, 255, 0);

	m_RedBrush = CreateSolidBrush(m_RedColor);
	m_GreenBrush = CreateSolidBrush(m_GreenColor);
}

BOOL COfficeActivatorDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	g_Mutex = CreateGlobalMutex();
	ASSERT(g_Mutex);

	CListCtrl* list = static_cast<CListCtrl*>(GetDlgItem(IDC_LIST_INSTALLED_LICENSE));
	ASSERT(list);
	InitializeListControl(list);
	UpdateLicenseList(list);

	CComboBox* pks = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_PRODUCT_KEY));
	ASSERT(pks);
	InitializeProductKeyList(pks);

	if (!IsUserAnAdmin()) {
		int buttons[] = { IDC_BUTTON_PATCH,IDC_BUTTON_RESTORE,IDC_BUTTON_INSTALL_PKEY,IDC_BUTTON_INSTALL_LIC,IDC_BUTTON_UNINSTALL_LIC,IDC_BUTTON_UNINSTALL_PKEY, IDC_BUTTON_SVC_MANAGE };
		for (int i = 0; i < sizeof(buttons) / sizeof(int); ++i) {
			CButton* btn = static_cast<CButton*>(GetDlgItem(buttons[i]));
			ASSERT(btn);

			Button_SetElevationRequiredState(btn->GetSafeHwnd(), TRUE);
		}
	}

	AutoDetectOffice();

	InitializeLicenseDirectoryMap();

	CComboBox* cbx = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_LICENSES_VERSION));
	ASSERT(cbx);
	InitializeLicenseFileControls(cbx);

#ifdef _WIN64
	SetWindowText(_T("Microsoft Office Activator [x64]"));
#else
	SetWindowText(_T("Microsoft Office Activator [x86]"));
#endif

	HWND hwnd = this->GetSafeHwnd();
	BOOL success = !!SetWindowLongPtr(hwnd, GWL_STYLE, GetWindowLongPtr(hwnd, GWL_STYLE) & ~WS_THICKFRAME);
	ASSERT(success);

	return TRUE;
}

BEGIN_MESSAGE_MAP(COfficeActivatorDlg, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON_BROWSE_OFFICE, &COfficeActivatorDlg::OnBnClickedButtonBrowseOffice)
	ON_BN_CLICKED(IDC_BUTTON_BROWSE_MSO, &COfficeActivatorDlg::OnBnClickedButtonBrowseMso)
	ON_BN_CLICKED(IDC_BUTTON_CHECK_PATCH_STATUS, &COfficeActivatorDlg::OnBnClickedButtonCheckPatchStatus)
	ON_BN_CLICKED(IDC_BUTTON_PATCH, &COfficeActivatorDlg::OnBnClickedButtonPatch)
	ON_BN_CLICKED(IDC_BUTTON_AUTO_DETECT, &COfficeActivatorDlg::OnBnClickedButtonAutoDetect)
	ON_BN_CLICKED(IDC_BUTTON_RESTORE, &COfficeActivatorDlg::OnBnClickedButtonRestore)
	ON_BN_CLICKED(IDC_BUTTON_UPDATE_LIC, &COfficeActivatorDlg::OnBnClickedButtonUpdateLic)
	ON_BN_CLICKED(IDC_BUTTON_UNINSTALL_PKEY, &COfficeActivatorDlg::OnBnClickedButtonUninstallPkey)
	ON_BN_CLICKED(IDC_BUTTON_INSTALL_PKEY, &COfficeActivatorDlg::OnBnClickedButtonInstallPkey)
	ON_CBN_SELCHANGE(IDC_COMBO_LICENSES_VERSION, &COfficeActivatorDlg::OnCbnSelchangeComboLicensesVersion)
	ON_BN_CLICKED(IDC_BUTTON_INSTALL_LIC, &COfficeActivatorDlg::OnBnClickedButtonInstallLic)
	ON_BN_CLICKED(IDC_BUTTON_UNINSTALL_LIC, &COfficeActivatorDlg::OnBnClickedButtonUninstallLic)
	
	ON_WM_CTLCOLOR()
END_MESSAGE_MAP()

HBRUSH COfficeActivatorDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH br = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	//if (pWnd->GetSafeHwnd() == this->GetSafeHwnd()) {
		pDC->SetBkColor(RGB(255, 255, 255));
		br = (HBRUSH)GetStockObject(WHITE_BRUSH);
	//}

	int nItem = pWnd->GetDlgCtrlID();
	if (m_PatchState != ~0) {
		if (nItem == IDC_EDIT_SPPCHOOK_PATCH_STATE) {
			if (m_PatchState & SPPCHOOK_NOT_FOUND) {
				br = m_RedBrush;
				pDC->SetBkColor(m_RedColor);
			}
			else {
				br = m_GreenBrush;
				pDC->SetBkColor(m_GreenColor);
			}
		}

		if (nItem == IDC_EDIT_MSO_PATCH_STATE) {
			if (m_PatchState & (INVALID_MSO_STATE | MSO_NOT_PATCHED)) {
				br = m_RedBrush;
				pDC->SetBkColor(m_RedColor);
			}
			else {
				br = m_GreenBrush;
				pDC->SetBkColor(m_GreenColor);
			}
		}
	}

	return br;
}

void COfficeActivatorDlg::OnBnClickedButtonBrowseOffice()
{
	CFolderPickerDialog* dlg = new CFolderPickerDialog();
	ASSERT(dlg);

	if (IDOK == dlg->DoModal()) {
		SetDlgItemText(IDC_EDIT_OFFICE_PATH, dlg->GetPathName());
		m_PatchState = ~0;
	}

	delete dlg;
}

void COfficeActivatorDlg::OnBnClickedButtonBrowseMso()
{
	CFileDialog* dlg = new CFileDialog(TRUE, _T(".dll"), _T("MSO.DLL"), OFN_OVERWRITEPROMPT, _T("DLL Files (*.dll)|*.dll"));
	ASSERT(dlg);

	if (IDOK == dlg->DoModal()) {
		if (dlg->GetFileName().CompareNoCase(_T("mso.dll")) != 0) {
			MessageBox(_T("Please select a valid MSO.DLL file!"), _T("Warning"), MB_OK);
		}
		else {
			SetDlgItemText(IDC_EDIT_MSO_PATH, dlg->GetPathName());
			m_PatchState = ~0;
		}
	}

	delete dlg;
}

void COfficeActivatorDlg::OnBnClickedButtonAutoDetect()
{
	AutoDetectOffice();
}

void COfficeActivatorDlg::OnBnClickedButtonCheckPatchStatus()
{
	WaitForSingleObject(g_Mutex, INFINITE);

	DWORD checkResult = 0;
	CString officePath, msoPath;
	if (GetDlgItemText(IDC_EDIT_OFFICE_PATH, officePath) == 0 ||
		GetDlgItemText(IDC_EDIT_MSO_PATH, msoPath) == 0) {
		MessageBox(_T("Please select the office installation path and the location of the MSO.DLL file."), _T("Info"), MB_OK);
		return;
	}

	CEdit* officePatch = static_cast<CEdit*>(GetDlgItem(IDC_EDIT_SPPCHOOK_PATCH_STATE));
	CEdit* msoPatch = static_cast<CEdit*>(GetDlgItem(IDC_EDIT_MSO_PATCH_STATE));
	ASSERT(officePatch);
	ASSERT(msoPatch);

	if (PathFileExists(officePath + _T("\\SppcHook.dll"))) {
		officePatch->SetWindowTextW(_T("SppcHook.dll exists."));
	}
	else {
		officePatch->SetWindowTextW(_T("SppcHook.dll not found."));
		checkResult |= SPPCHOOK_NOT_FOUND;
	}

	if (PathFileExists(msoPath)) {

		switch (IsMsoPatched(msoPath)) {
		case 0:
			//
			// Unpatched.
			//

			if (VerifyEmbeddedSignature(msoPath)) {
				msoPatch->SetWindowTextW(_T("MSO.DLL is not patched."));
				checkResult |= MSO_NOT_PATCHED;
			}
			else {
				msoPatch->SetWindowTextW(_T("The status of MSO.DLL is invalid."));
				checkResult |= INVALID_MSO_STATE;
			}

			break;

		case 1:
			//
			// Patched.
			//

			msoPatch->SetWindowTextW(_T("MSO.DLL has been patched."));
			break;

		case -1:
			//
			// Fail
			//

			msoPatch->SetWindowTextW(_T("The status of MSO.DLL is invalid."));
			checkResult |= INVALID_MSO_STATE;
			break;
		}

	}
	else {
		msoPatch->SetWindowTextW(_T("MSO.DLL not found."));
		checkResult |= INVALID_MSO_STATE;
	}

	m_PatchState = checkResult;

	officePatch->Invalidate();
	msoPatch->Invalidate();

	ReleaseMutex(g_Mutex);
}

void COfficeActivatorDlg::OnBnClickedButtonPatchInternal()
{
	if (m_PatchState == ~0) {
		MessageBox(_T("Please check patch state first!"), _T("Error"));
		return;
	}

	if ((m_PatchState & INVALID_MSO_STATE) == INVALID_MSO_STATE) {
		MessageBox(_T("The status of MSO.DLL is invalid."), _T("Warning"));
		return;
	}

	if (m_PatchState == 0) {
		MessageBox(_T("The patch has already been applied."), _T("Info"));
		return;
	}

	if (!PathFileExists(SPPC_NAME)) {
		MessageBox(_T("SppcHook[64/32].dll is missing."), _T("Info"));
		return;
	}

	do {
		if (m_PatchState & SPPCHOOK_NOT_FOUND) {
			CString officePath;
			GetDlgItemText(IDC_EDIT_OFFICE_PATH, officePath);

			if (!CopyFile(SPPC_NAME, officePath + _T("\\SppcHook.dll"), TRUE)) {
				CString msg;
				msg.Format(_T("Copy SppcHook.dll failed: %d."), GetLastError());

				MessageBox(msg, _T("Error"));
				break;
			}
		}

		if (m_PatchState & MSO_NOT_PATCHED) {
			CString msoPath;
			GetDlgItemText(IDC_EDIT_MSO_PATH, msoPath);

			DeleteFile(msoPath + _T(".bak"));

			if (!MoveFileW(msoPath, msoPath + _T(".bak"))) {
				CString msg;
				msg.Format(_T("Backup MSO.DLL failed: %d."), GetLastError());

				MessageBox(msg, _T("Error"));
				break;
			}

			switch (PatchMsoFile(msoPath + _T(".bak"), msoPath)) {
			case 0:
				MessageBox(_T("The patch was applied successfully."), _T("Info"));
				break;

			case 1:
				MessageBox(_T("Patch failed: Cannot read source file."), _T("Info"));
				break;

			case 2:
				MessageBox(_T("Patch failed: Cannot create new file."), _T("Info"));
				break;

			case 3:
				MessageBox(_T("Patch failed: Invalid file format."), _T("Info"));
				break;

			case 4:
			case 5:
				MessageBox(_T("Patch failed: Unsupported file."), _T("Warning"));
				break;

			case 6:
				MessageBox(_T("Patch failed: Insufficient memory resource."), _T("Warning"));
				break;

			case 7:
				MessageBox(_T("Patch failed: Write file failed."), _T("Warning"));
				break;

			case -1:
				MessageBox(_T("Patch failed."), _T("Warning"));
				break;
			}
		}

	} while (false);
}

void COfficeActivatorDlg::OnBnClickedButtonPatch() {
	if (MessageBox(_T("DO YOU REALLY WANT TO APPLY PATCH?"), _T("WARNING"), MB_YESNOCANCEL | MB_ICONWARNING) != IDYES) {
		return;
	}

	WaitForSingleObject(g_Mutex, INFINITE);
	OnBnClickedButtonCheckPatchStatus();

	__try {
		OnBnClickedButtonPatchInternal();
	}
	__finally {
		OnBnClickedButtonCheckPatchStatus();
		ReleaseMutex(g_Mutex);
	}
}

void COfficeActivatorDlg::OnBnClickedButtonRestoreInternal()
{
	if (m_PatchState == ~0) {
		MessageBox(_T("Please check patch state first!"), _T("Error"));
		return;
	}

	if ((m_PatchState & INVALID_MSO_STATE) == INVALID_MSO_STATE) {
		MessageBox(_T("The status of MSO.DLL is invalid."), _T("Warning"));
		return;
	}

	if ((m_PatchState & SPPCHOOK_NOT_FOUND) == 0) {
		CString officePath;

		GetDlgItemText(IDC_EDIT_OFFICE_PATH, officePath);

		if (!DeleteFile(officePath + _T("\\SppcHook.dll"))) {
			CString msg;
			msg.Format(_T("Delete SppcHook.dll failed: %d."), GetLastError());

			MessageBox(msg, _T("Error"));
			return;
		}
	}

	if (!(m_PatchState & MSO_NOT_PATCHED)) {
		CString msoPath;

		GetDlgItemText(IDC_EDIT_MSO_PATH, msoPath);

		if (!PathFileExists(msoPath + _T(".bak"))) {
			MessageBox(_T("MSO.DLL.bak not found."), _T("Error"));
			return;
		}

		if (!DeleteFile(msoPath)) {
			CString msg;
			msg.Format(_T("Delete MSO.DLL failed: %d."), GetLastError());

			MessageBox(msg, _T("Error"));
			return;
		}

		if (!MoveFile(msoPath + _T(".bak"), msoPath)) {
			CString msg;
			msg.Format(_T("Restore MSO.DLL failed: %d."), GetLastError());

			MessageBox(msg, _T("Error"));
			return;
		}
	}
}

void COfficeActivatorDlg::OnBnClickedButtonRestore() {
	if (MessageBox(_T("DO YOU REALLY WANT TO RESTORE PATCH?\n(Maybe you need disable or delete OfficeActivatorSvc)"), _T("WARNING"), MB_YESNOCANCEL | MB_ICONWARNING) != IDYES) {
		return;
	}

	WaitForSingleObject(g_Mutex, INFINITE);
	OnBnClickedButtonCheckPatchStatus();

	__try {
		OnBnClickedButtonRestoreInternal();
	}
	__finally {
		//
		// Update patch status.
		//
		OnBnClickedButtonCheckPatchStatus();
		ReleaseMutex(g_Mutex);
	}
}

void COfficeActivatorDlg::OnBnClickedButtonUpdateLic()
{
	CListCtrl* list = static_cast<CListCtrl*>(GetDlgItem(IDC_LIST_INSTALLED_LICENSE));
	ASSERT(list);

	UpdateLicenseList(list);
}

void COfficeActivatorDlg::OnBnClickedButtonUninstallPkey()
{
	if (MessageBox(_T("DO YOU REALLY WANT TO UNINSTALL SELECTED PKEY(S)?"), _T("WARNING"), MB_YESNOCANCEL | MB_ICONWARNING) != IDYES) {
		return;
	}

	CListCtrl* list = static_cast<CListCtrl*>(GetDlgItem(IDC_LIST_INSTALLED_LICENSE));
	ASSERT(list);

	DWORD checkCount = 0;
	DWORD successCount = 0;
	std::vector<CString>keys;

	for (int i = 0; i < list->GetItemCount(); ++i) {
		if (list->GetCheck(i)) {
			CString pk = list->GetItemText(i, 3);
			++checkCount;

			if (pk.GetLength() != 0) {
				keys.push_back(pk);
			}

		}
	}

	if (!keys.empty()) {
		LPCWSTR* buf = new LPCWSTR[keys.size()];
		for (DWORD i = 0; i < keys.size(); ++i) {
			buf[i] = keys[i];
		}

		successCount = UninstallProductKey((DWORD)keys.size(), buf);
		delete[]buf;
	}

	CString msg;
	msg.Format(_T("Select: %d, Success: %d, Skip: %d"), checkCount, successCount, checkCount - (DWORD)keys.size());
	MessageBox(msg, _T("Info"));

	OnBnClickedButtonUpdateLic();
}

void COfficeActivatorDlg::OnBnClickedButtonInstallPkey()
{
	CComboBox* cbx = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_PRODUCT_KEY));
	ASSERT(cbx);

	CString pk;
	int sel = cbx->GetCurSel();
	if (sel >= 0) {
		pk = (LPCTSTR)cbx->GetItemDataPtr(sel);
	}
	else {
		GetDlgItemText(IDC_COMBO_PRODUCT_KEY, pk);
	}

	if (pk.GetLength() == 0) {
		MessageBox(_T("Please input product key."), _T("Warning"));
		return;
	}

	if (MessageBox(_T("DO YOU REALLY WANT TO INSTALL THIS PKEY?"), _T("WARNING"), MB_YESNOCANCEL | MB_ICONWARNING) != IDYES) {
		return;
	}

	HRESULT hr = InstallProductKey(pk);
	if (SUCCEEDED(hr)) {
		MessageBox(_T("Success."), _T("Info"));
	}
	else {
		MessageBox(_T("Fail: ") + SLErrorToString(hr), _T("Info"));
	}

	OnBnClickedButtonUpdateLic();
}

void COfficeActivatorDlg::OnCbnSelchangeComboLicensesVersion()
{
	CComboBox* cbx = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_LICENSES_VERSION));
	ASSERT(cbx);

	CComboBox* cbx2 = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_LICENSES));
	ASSERT(cbx2);
	cbx2->ResetContent();

	int sel = cbx->GetCurSel();
	ASSERT(sel >= 0);

	auto map = g_LicenseList[sel].second;
	for (auto iter = map.cbegin(); iter != map.cend(); ++iter) {
		cbx2->AddString(iter->first);
	}
}

void COfficeActivatorDlg::OnBnClickedButtonInstallLic()
{
	CComboBox* cbx = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_LICENSES_VERSION));
	ASSERT(cbx);
	CComboBox* cbx2 = static_cast<CComboBox*>(GetDlgItem(IDC_COMBO_LICENSES));
	ASSERT(cbx2);

	int sel = cbx2->GetCurSel();
	if (sel == -1) {
		MessageBox(_T("Please choose a license to install."), _T("Warning"));
		return;
	}

	if (MessageBox(_T("DO YOU REALLY WANT TO INSTALL SELECTED LICENSE?"), _T("WARNING"), MB_YESNOCANCEL | MB_ICONWARNING) != IDYES) {
		return;
	}

	int sel2 = cbx->GetCurSel();
	ASSERT(sel2 != -1);

	auto pair = g_LicenseList[sel2];
	CString path;
	path.Format(_T("licenses\\%d\\"), pair.first);

	CString name;
	GetDlgItemText(IDC_COMBO_LICENSES, name);

	auto map = pair.second;
	auto iter = map.find(name);
	if (iter == map.end()) {
		ASSERT(FALSE);
	}

	BOOL success = TRUE;
	for (DWORD i = 0; i < sizeof(g_TypeTable) / sizeof(LPCTSTR); ++i) {
		if (iter->second.dwFlags & (1 << i)) {
			CString fileName = path + name + _T('-') + g_TypeTable[i];
			HRESULT hr = InstallLicense(fileName);
			
			if (!SUCCEEDED(hr)) {
				MessageBox(_T("Install license failed: ") + SLErrorToString(hr), _T("Error"));
				success = FALSE;
				break;
			}

		}
	}

	if (success) {
		MessageBox(_T("Install license success"), _T("Info"));
		success = FALSE;
	}

	OnBnClickedButtonUpdateLic();
}

void COfficeActivatorDlg::OnBnClickedButtonUninstallLic()
{
	if (MessageBox(_T("DO YOU REALLY WANT TO UNINSTALL SELECTED LICENSE(S)?"), _T("WARNING"), MB_YESNOCANCEL | MB_ICONWARNING) != IDYES) {
		return;
	}

	CListCtrl* list = static_cast<CListCtrl*>(GetDlgItem(IDC_LIST_INSTALLED_LICENSE));
	ASSERT(list);

	DWORD checkCount = 0;
	DWORD successCount = 0;

	for (int i = 0; i < list->GetItemCount(); ++i) {
		if (list->GetCheck(i)) {
			CString name = list->GetItemText(i, 0);
			++checkCount;

			if (name.GetLength() != 0) {
				CString pn = name.Left(9);
				WORD index = 0;

				if (pn == _T("Office 15")) {
					index = 2013;
				}
				else if (pn == _T("Office 16")) {
					index = 2016;
				}
				else if (pn == _T("Office 19")) {
					index = 2019;
				}
				else if (pn == _T("Office 21")) {
					index = 2021;
				}
				else {
					index = 0;
				}

				if (index == 0) {
					MessageBox(_T("The license file associated with the license you selected could not be found."), _T("Info"));
				}
				else {
					for (auto i = g_LicenseList.cbegin(); i != g_LicenseList.cend(); ++i) {
						if (i->first == index) {
							pn = name.Mid(11);
							int pos = pn.Find(_T(' '));
							if (pos == -1) {
								ASSERT(FALSE);
							}
							else {
								pn = pn.Left(pos);
							}

							auto iter = i->second.find(pn);
							if (iter == i->second.end()) {
								MessageBox(_T("The license file associated with the license you selected could not be found."), _T("Info"));
							}
							else {
								CString path;
								BOOL success = FALSE;
								path.Format(_T("licenses\\%d\\"), index);

								for (DWORD i = 0; i < sizeof(g_TypeTable) / sizeof(LPCTSTR); ++i) {
									if (iter->second.dwFlags & (1 << i)) {
										HRESULT hr = UninstallLicense(path + pn + _T("-") + g_TypeTable[i]);
										if (SUCCEEDED(hr))success = TRUE;
									}
								}

								if (success)++successCount;
							}

							break;
						}
					}
				}

			}

		}
	}

	CString msg;
	msg.Format(_T("Select: %d, Success: %d"), checkCount, successCount);
	MessageBox(msg, _T("Info"));

	OnBnClickedButtonUpdateLic();
}
