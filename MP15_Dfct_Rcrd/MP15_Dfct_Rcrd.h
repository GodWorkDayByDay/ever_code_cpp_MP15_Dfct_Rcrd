
// MP15_Dfct_Rcrd.h : PROJECT_NAME 應用程式的主要標頭檔
//

#pragma once

#ifndef __AFXWIN_H__
	#error "對 PCH 包含此檔案前先包含 'stdafx.h'"
#endif

#include "resource.h"		// 主要符號


// CMP15_Dfct_RcrdApp:
// 請參閱實作此類別的 MP15_Dfct_Rcrd.cpp
//

class CMP15_Dfct_RcrdApp : public CWinApp
{
public:
	CMP15_Dfct_RcrdApp();

// 覆寫
public:
	virtual BOOL InitInstance();

// 程式碼實作

	DECLARE_MESSAGE_MAP()
};

extern CMP15_Dfct_RcrdApp theApp;