#pragma once


// CServiceManageDlg dialog

class CServiceManageDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CServiceManageDlg)

	LPCTSTR lpBinPath;

	SC_HANDLE hSCM;

	CStatic* label;
	CButton* btnCreate;
	CButton* btnStart;
	CButton* btnStop;
	CButton* btnDelete;
	CButton* btnDisable;
	CButton* btnEnable;

	VOID UpdateServiceStatus(SC_HANDLE hSvc);

	SC_HANDLE TryOpenService();

public:
	CServiceManageDlg(LPCTSTR lpBinPath, CWnd* pParent = nullptr);   // standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_SVC };
#endif

protected:
	virtual BOOL OnInitDialog();
	void OnClose();

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonCreate();
	afx_msg void OnBnClickedButtonStart();
	afx_msg void OnBnClickedButtonStop();
	afx_msg void OnBnClickedButtonDelete();
	afx_msg void OnBnClickedButtonDisableSvc();
	afx_msg void OnBnClickedButtonEnableSvc();
	afx_msg void OnBnClickedButtonRefreshSvc();
};
