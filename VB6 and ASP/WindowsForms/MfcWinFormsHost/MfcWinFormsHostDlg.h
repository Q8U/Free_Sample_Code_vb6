// MfcWinFormsHostDlg.h : header file
//

#pragma once

// DONE: Some helper classes
#include "WinFormsControlHelpers.h"


// CMfcWinFormsHostDlg dialog
class CMfcWinFormsHostDlg : public CDialog
{
  // Construction
public:
  CMfcWinFormsHostDlg(CWnd* pParent = NULL);	// standard constructor

  // Dialog Data
  enum { IDD = IDD_MFCWINFORMSHOST_DIALOG };

protected:
  virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


  // Implementation
protected:
  HICON m_hIcon;

  // Generated message map functions
  virtual BOOL OnInitDialog();
  afx_msg void OnPaint();
  afx_msg HCURSOR OnQueryDragIcon();
  DECLARE_MESSAGE_MAP()
public:

  // DONE: Host WinForms controls as COM controls
  virtual BOOL CreateControlSite(COleControlContainer* pContainer, COleControlSite** ppSite, UINT nID, REFCLSID clsid);

private:
  // DONE: A wrapper for a WinForms control
  CWinFormsControlWnd m_wndWinFormsCalendar;
};












