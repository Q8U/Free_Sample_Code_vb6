// anotherMfcTest.h : main header file for the anotherMfcTest application
//
#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols


// CanotherMfcTestApp:
// See anotherMfcTest.cpp for the implementation of this class
//

class CanotherMfcTestApp : public CWinApp
{
public:
	CanotherMfcTestApp();


// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CanotherMfcTestApp theApp;