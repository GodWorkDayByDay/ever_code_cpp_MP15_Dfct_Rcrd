
// MP15_Dfct_RcrdDlg.h : 標頭檔
//

#pragma once
#include "stdafx.h"
#include "ADO_Conn.h"
#include "List_Row_Data.h"
#include <vector>
#include <algorithm>
#include "Ctx_Db.h"

#include "comdef.h"
#include <cmath>


#include "CApplication.h"
#include "CRange.h"
#include "CWorkbooks.h"
#include "CWorksheets.h"
#include "CWorkbook.h"
#include "CWorksheet.h"

//#include <windows.h>

//#import "c:\Program Files\Common Files\System\ado\msado15.dll" no_namespace rename("EOF","adoEOF") rename("BOF","adoBOF")
 
// CMP15_Dfct_RcrdDlg 對話方塊
class CMP15_Dfct_RcrdDlg : public CDialogEx
{
// 建構
public:
	CMP15_Dfct_RcrdDlg(CWnd* pParent = NULL);	// 標準建構函式

// 對話方塊資料
	enum { IDD = IDD_MP15_DFCT_RCRD_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支援

	ADOConn Db_Ctx;

// 程式碼實作
protected:
	HICON m_hIcon;

	// 產生的訊息對應函式
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnRclickListDfct(NMHDR *pNMHDR, LRESULT *pResult);
	CListCtrl m_CList;
	afx_msg void OnBnClickedButtonQry();

	std::vector<CWnd *> HCtrSw;
	std::vector<CWnd *> ::iterator ItrHCtrSw;

	std::vector<CString> Col_Name;
	std::vector<CString>::iterator ItrCol;

	std::vector<List_Row_Data> Lst_Data_Dfct;
	std::vector<List_Row_Data> Lst_Data_Pass;

	std::vector<List_Row_Data> ToDay_Lst_Data_Dfct;
	std::vector<List_Row_Data> ToDay_Lst_Data_Pass;

	std::vector<List_Row_Data> Lst_Data_Qry;

	CString m_Box_Numb;
	CString m_SN_Numb;
//	CString m_Dfct_1;
//	CString m_Dfct_2;
	CString m_Line;
	afx_msg void OnBnClickedButtonDft();
	afx_msg void OnBnClickedButtonCfm();
	afx_msg void OnBnClickedButtonClr();
	// // SQL SERVER , Default is 113
	ADOConn Data_Base_Ctrl;
	
	afx_msg void OnIdrPass();

public:
	//CEdit m_SubItemEdit;
	afx_msg void OnDblclkListDfct(NMHDR *pNMHDR, LRESULT *pResult);
	CEdit m_SubItemEdit;
	CString m_Ipt_Pss_Info;
	afx_msg void OnKillfocusEditHideIpt();
	int Row_Edit_Idx;
	int Col_Edit_Index;
	int GetDfct(void);
	int GetPss(void);
	int Lst_Data_Dfct_Show_Lst(void);
	int Lst_Data_Pss_Show_Lst(void);
	int Shw_Pss_Lst_Flg;
	afx_msg void OnBnClickedButtonStartQry();
//	CString m_Dfct_Cnt;
//	CString m_Pss_Cnt;
	int m_Dfct_Cnt;
	int m_Pss_Cnt;
	int m_Qry_Flg;
	COleDateTime m_Start_Time;
	COleDateTime m_End_Time;
	int m_Add_Cnt;
	CString m_Ord_Numb;
	afx_msg void OnBnClickedButtonExcel();
	int Lst_Mry_Excel_Flg;
	CString m_Dfct_1;
//	CComboBox m_Dfct_2;
	CString m_Dfct_2;
	CString m_Lst_State;
	int m__Today_Pass_Cnt;
	int m_Today_Dfct_Cnt;
	int Get_Today_Dfct(void);
	int Get_Today_Pass(void);
	afx_msg void OnCustomdrawListDfct(NMHDR *pNMHDR, LRESULT *pResult);
	int Check_Box_Numb(void);
};
