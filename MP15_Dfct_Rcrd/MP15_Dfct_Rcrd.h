
// MP15_Dfct_Rcrd.h : PROJECT_NAME ���ε{�����D�n���Y��
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�� PCH �]�t���ɮ׫e���]�t 'stdafx.h'"
#endif

#include "resource.h"		// �D�n�Ÿ�


// CMP15_Dfct_RcrdApp:
// �аѾ\��@�����O�� MP15_Dfct_Rcrd.cpp
//

class CMP15_Dfct_RcrdApp : public CWinApp
{
public:
	CMP15_Dfct_RcrdApp();

// �мg
public:
	virtual BOOL InitInstance();

// �{���X��@

	DECLARE_MESSAGE_MAP()
};

extern CMP15_Dfct_RcrdApp theApp;