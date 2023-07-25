// CServiceManageDlg.cpp : implementation file
//

#include "pch.h"
#include "OfficeActivator.h"
#include <winsvc.h>
#include "ServiceManageDlg.h"
#include "afxdialogex.h"
#include "Service.h"


// CServiceManageDlg dialog

IMPLEMENT_DYNAMIC(CServiceManageDlg, CDialogEx)

CServiceManageDlg::CServiceManageDlg(LPCTSTR lpBinPath, CWnd* pParent)
	: CDialogEx(IDD_DIALOG_SVC, pParent), lpBinPath(lpBinPath) {
	hSCM = nullptr;
	label = nullptr;
	btnCreate = nullptr;
	btnStart = nullptr;
	btnStop = nullptr;
	btnDelete = nullptr;
	btnDisable = nullptr;
	btnEnable = nullptr;
}

BOOL CServiceManageDlg::OnInitDialog() {
	CDialogEx::OnInitDialog();

	label = static_cast<CStatic*>(GetDlgItem(IDC_STATIC_STATUS));
	btnCreate = static_cast<CButton*>(GetDlgItem(IDC_BUTTON_CREATE_SVC));
	btnStart = static_cast<CButton*>(GetDlgItem(IDC_BUTTON_START_SVC));
	btnStop = static_cast<CButton*>(GetDlgItem(IDC_BUTTON_STOP_SVC));
	btnDelete = static_cast<CButton*>(GetDlgItem(IDC_BUTTON_DELETE_SVC));
	btnDisable = static_cast<CButton*>(GetDlgItem(IDC_BUTTON_DISABLE_SVC));
	btnEnable = static_cast<CButton*>(GetDlgItem(IDC_BUTTON_ENABLE_SVC));
	ASSERT(label);
	ASSERT(btnCreate);
	ASSERT(btnStart);
	ASSERT(btnStop);
	ASSERT(btnDelete);
	ASSERT(btnDisable);
	ASSERT(btnEnable);

	hSCM = OpenSCManager(nullptr, nullptr, SC_MANAGER_ALL_ACCESS);
	if (!hSCM) {
		MessageBox(_T("Can not open ServiceManager."), _T("Error"));
		EndDialog(-1);
		return TRUE;
	}

	SC_HANDLE hSvc = OpenService(hSCM, SVCNAME, SERVICE_ALL_ACCESS);
	if (hSvc) {
		UpdateServiceStatus(hSvc);
		CloseServiceHandle(hSvc);
	}
	else {
		if (GetLastError() != ERROR_SERVICE_DOES_NOT_EXIST) {
			MessageBox(_T("Can not open Service."), _T("Error"));
			EndDialog(-1);
			return TRUE;
		}

		label->SetWindowText(_T("Non-existent"));
		btnCreate->EnableWindow();
		btnDelete->EnableWindow(FALSE);
		btnStart->EnableWindow(FALSE);
		btnStop->EnableWindow(FALSE);
	}

	return TRUE;
}

SC_HANDLE CServiceManageDlg::TryOpenService() {
	SC_HANDLE hSvc = OpenService(hSCM, SVCNAME, SERVICE_ALL_ACCESS);
	if (!hSvc) {
		label->SetWindowText(_T("Non-existent"));
		btnCreate->EnableWindow();
		btnDelete->EnableWindow(FALSE);
		btnStart->EnableWindow(FALSE);
		btnStop->EnableWindow(FALSE);
	}

	return hSvc;
}

BEGIN_MESSAGE_MAP(CServiceManageDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTTON_CREATE_SVC, &CServiceManageDlg::OnBnClickedButtonCreate)
	ON_BN_CLICKED(IDC_BUTTON_START_SVC, &CServiceManageDlg::OnBnClickedButtonStart)
	ON_BN_CLICKED(IDC_BUTTON_STOP_SVC, &CServiceManageDlg::OnBnClickedButtonStop)
	ON_BN_CLICKED(IDC_BUTTON_DELETE_SVC, &CServiceManageDlg::OnBnClickedButtonDelete)
	ON_BN_CLICKED(IDC_BUTTON_DISABLE_SVC, &CServiceManageDlg::OnBnClickedButtonDisableSvc)
	ON_BN_CLICKED(IDC_BUTTON_ENABLE_SVC, &CServiceManageDlg::OnBnClickedButtonEnableSvc)
	ON_BN_CLICKED(IDC_BUTTON_REFRESH_SVC, &CServiceManageDlg::OnBnClickedButtonRefreshSvc)
END_MESSAGE_MAP()

VOID CServiceManageDlg::UpdateServiceStatus(SC_HANDLE hSvc) {
	btnCreate->EnableWindow(FALSE);
	btnDelete->EnableWindow();

	SERVICE_STATUS ServiceStatus;
	if (QueryServiceStatus(hSvc, &ServiceStatus)) {
		switch (ServiceStatus.dwCurrentState) {
		case SERVICE_STOPPED:
			btnStart->EnableWindow();
			btnStop->EnableWindow(FALSE);
			label->SetWindowText(_T("Stopped"));
			break;

		case SERVICE_START_PENDING:
			btnStart->EnableWindow();
			btnStop->EnableWindow();
			label->SetWindowText(_T("Starting"));
			break;

		case SERVICE_STOP_PENDING:
			btnStart->EnableWindow();
			btnStop->EnableWindow();
			label->SetWindowText(_T("Stopping"));
			break;

		case SERVICE_RUNNING:
			btnStart->EnableWindow(FALSE);
			btnStop->EnableWindow();
			label->SetWindowText(_T("Running"));
			break;

		default:
			btnStart->EnableWindow(FALSE);
			btnStop->EnableWindow(FALSE);
			label->SetWindowText(_T("Unkonwn Status"));
			break;
		}
	}
	else {
		btnStart->EnableWindow(FALSE);
		btnStop->EnableWindow(FALSE);
		label->SetWindowText(_T("Unkonwn Status"));
	}
}

// CServiceManageDlg message handlers

void CServiceManageDlg::OnClose() {
	if (hSCM) {
		CloseServiceHandle(hSCM);
		hSCM = nullptr;
	}

	CDialog::OnClose();
}

void CServiceManageDlg::OnBnClickedButtonCreate()
{
	SERVICE_DELAYED_AUTO_START_INFO sdasi;
	sdasi.fDelayedAutostart = TRUE;

	SC_HANDLE hSvc = CreateService(
		hSCM,
		SVCNAME,
		SVCDNAME,
		SERVICE_ALL_ACCESS,
		SERVICE_WIN32_OWN_PROCESS,
		SERVICE_AUTO_START,
		SERVICE_ERROR_NORMAL,
		lpBinPath,
		nullptr,
		nullptr,
		nullptr,
		nullptr,
		nullptr
	);
	if (hSvc &&
		ChangeServiceConfig2(
			hSvc,
			SERVICE_CONFIG_DELAYED_AUTO_START_INFO,
			&sdasi)) {

		UpdateServiceStatus(hSvc);
		CloseServiceHandle(hSvc);

		MessageBox(_T("Create Service Successfull"), _T("Info"));
	}
	else {
		if (GetLastError() == ERROR_SERVICE_EXISTS) {
			hSvc = OpenService(hSCM, SVCNAME, SERVICE_ALL_ACCESS);
			if (hSvc) {
				UpdateServiceStatus(hSvc);
				CloseServiceHandle(hSvc);

				MessageBox(_T("Service Already Exists"), _T("Info"));
			}
			else {
				MessageBox(_T("Create Service Failed"), _T("Info"));
			}
		}
		else {
			MessageBox(_T("Create Service Failed"), _T("Info"));
		}
	}
}

void CServiceManageDlg::OnBnClickedButtonStart()
{
	SC_HANDLE hSvc = TryOpenService();
	if (hSvc) {
		if (StartService(hSvc, 0, nullptr)) {
			MessageBox(_T("Service start successful"), _T("Error"));
		}
		else {
			MessageBox(_T("Service start failed"), _T("Error"));
		}

		UpdateServiceStatus(hSvc);
		CloseServiceHandle(hSvc);
	}
	else {
		MessageBox(_T("Service not exists"), _T("Error"));
	}
}

void CServiceManageDlg::OnBnClickedButtonStop()
{
	SC_HANDLE hSvc = TryOpenService();
	if (hSvc) {
		SERVICE_STATUS status;
		if (ControlService(hSvc, SERVICE_CONTROL_STOP, &status)) {
			MessageBox(_T("Service stop successful"), _T("Error"));
		}
		else {
			MessageBox(_T("Service stop failed"), _T("Error"));
		}

		UpdateServiceStatus(hSvc);
		CloseServiceHandle(hSvc);
	}
	else {
		MessageBox(_T("Service not exists"), _T("Error"));
	}
}

void CServiceManageDlg::OnBnClickedButtonDelete()
{
	SC_HANDLE hSvc = TryOpenService();
	if (hSvc) {
		if (DeleteService(hSvc)) {
			MessageBox(_T("Service delete successful"), _T("Error"));
		}
		else {
			MessageBox(_T("Service delete failed"), _T("Error"));
		}

		UpdateServiceStatus(hSvc);
		CloseServiceHandle(hSvc);
	}
	else {
		MessageBox(_T("Service not exists"), _T("Error"));
	}
}

void CServiceManageDlg::OnBnClickedButtonDisableSvc()
{
	// TODO: Add your control notification handler code here
}

void CServiceManageDlg::OnBnClickedButtonEnableSvc()
{
	// TODO: Add your control notification handler code here
}

void CServiceManageDlg::OnBnClickedButtonRefreshSvc()
{
	SC_HANDLE hSvc = TryOpenService();
	if (hSvc) {
		UpdateServiceStatus(hSvc);
		CloseServiceHandle(hSvc);
	}
}
