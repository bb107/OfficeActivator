
// OfficeActivatorDlg.h : header file
//

#pragma once


// COfficeActivatorDlg dialog
class COfficeActivatorDlg : public CDialogEx
{
// Construction
public:
	COfficeActivatorDlg(CWnd* pParent = nullptr);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_OFFICEACTIVATOR_DIALOG };
#endif

private:
	DWORD m_PatchState;

	DWORD m_RedColor;
	HBRUSH m_RedBrush;

	DWORD m_GreenColor;
	HBRUSH m_GreenBrush;

	VOID AutoDetectOffice();

// Implementation
protected:
	HICON m_hIcon;

	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()

	// Generated message map functions
public:
	afx_msg void OnBnClickedButtonBrowseOffice();
	afx_msg void OnBnClickedButtonBrowseMso();
	afx_msg void OnBnClickedButtonCheckPatchStatus();
	afx_msg void OnBnClickedButtonPatch();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedButtonAutoDetect();
	afx_msg void OnBnClickedButtonRestore();
	afx_msg void OnBnClickedButtonUpdateLic();
	afx_msg void OnBnClickedButtonUninstallPkey();
	afx_msg void OnBnClickedButtonInstallPkey();
	afx_msg void OnCbnSelchangeComboLicensesVersion();
	afx_msg void OnBnClickedButtonInstallLic();
	afx_msg void OnBnClickedButtonUninstallLic();
};
