
// MP15_Dfct_RcrdDlg.cpp : 實作檔
//

#include "stdafx.h"
#include "MP15_Dfct_Rcrd.h"
#include "MP15_Dfct_RcrdDlg.h"
#include "afxdialogex.h"
#include <iostream>
using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CMP15_Dfct_RcrdDlg 對話方塊




CMP15_Dfct_RcrdDlg::CMP15_Dfct_RcrdDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CMP15_Dfct_RcrdDlg::IDD, pParent)
	, Row_Edit_Idx(0)
	, Col_Edit_Index(6)
	, Shw_Pss_Lst_Flg(0)
	, Lst_Mry_Excel_Flg(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_Box_Numb = _T("");
	m_SN_Numb = _T("");
	//  m_Dfct_1 = _T("");
	//  m_Dfct_2 = _T("");
	m_Line = _T("");

	/*Vct_MP_Main_Tab
	Ctx_Db_Ctrl*/
	m_Ipt_Pss_Info = _T("");
	//  m_Dfct_Cnt = _T("");
	//  m_Pss_Cnt = _T("");
	m_Dfct_Cnt = 0;
	m_Pss_Cnt = 0;
	m_Qry_Flg = 0;
	m_Start_Time = COleDateTime::GetCurrentTime();
	m_End_Time = COleDateTime::GetCurrentTime();
	m_Add_Cnt = 0;
	m_Ord_Numb = _T("");
	m_Dfct_1 = _T("");
	m_Dfct_2 = _T("");
	m_Lst_State = _T("");
	m__Today_Pass_Cnt = 0;
	m_Today_Dfct_Cnt = 0;
}

void CMP15_Dfct_RcrdDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST__Dfct, m_CList);
	DDX_Text(pDX, IDC_EDIT_Box_Numb, m_Box_Numb);
	DDX_Text(pDX, IDC_EDIT_SN_Numb, m_SN_Numb);
	//  DDX_Text(pDX, IDC_EDIT_Dfct_1, m_Dfct_1);
	//  DDX_Text(pDX, IDC_EDIT_Dfct_2, m_Dfct_2);
	DDX_CBString(pDX, IDC_COMBO_Line, m_Line);
	DDX_Control(pDX, IDC_EDIT_Hide_Ipt, m_SubItemEdit);
	DDX_Text(pDX, IDC_EDIT_Hide_Ipt, m_Ipt_Pss_Info);
	//  DDX_Text(pDX, IDC_EDIT_Tol_Cnts, m_Dfct_Cnt);
	//  DDX_Text(pDX, IDC_EDIT_Pass_Cnts, m_Pss_Cnt);
	DDX_Text(pDX, IDC_EDIT_Tol_Cnts, m_Dfct_Cnt);
	DDX_Text(pDX, IDC_EDIT_Pass_Cnts, m_Pss_Cnt);
	DDX_Radio(pDX, IDC_RADIO_By_Time, m_Qry_Flg);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER_Str, m_Start_Time);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER_End, m_End_Time);
	DDX_Text(pDX, IDC_EDIT_Add_Cnt, m_Add_Cnt);
	DDX_Text(pDX, IDC_EDIT_Ord_Numb, m_Ord_Numb);
	DDX_CBString(pDX, IDC_COMBO_Dfct_1, m_Dfct_1);
	//  DDX_Control(pDX, IDC_COMBO_Dfct_2, m_Dfct_2);
	DDX_CBString(pDX, IDC_COMBO_Dfct_2, m_Dfct_2);
	DDX_Text(pDX, IDC_EDIT_Lst_State, m_Lst_State);
	DDX_Text(pDX, IDC_EDIT_Today_Pass_Cnt, m__Today_Pass_Cnt);
	DDX_Text(pDX, IDC_EDIT_Today_Dfct_Cnt, m_Today_Dfct_Cnt);
}

BEGIN_MESSAGE_MAP(CMP15_Dfct_RcrdDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_NOTIFY(NM_RCLICK, IDC_LIST__Dfct, &CMP15_Dfct_RcrdDlg::OnRclickListDfct)
	ON_BN_CLICKED(IDC_BUTTON_Qry, &CMP15_Dfct_RcrdDlg::OnBnClickedButtonQry)
	ON_BN_CLICKED(IDC_BUTTON_Dft, &CMP15_Dfct_RcrdDlg::OnBnClickedButtonDft)
	ON_BN_CLICKED(IDC_BUTTON_Cfm, &CMP15_Dfct_RcrdDlg::OnBnClickedButtonCfm)
	ON_BN_CLICKED(IDC_BUTTON_Clr, &CMP15_Dfct_RcrdDlg::OnBnClickedButtonClr)
	ON_COMMAND(ID_IDR_PASS, &CMP15_Dfct_RcrdDlg::OnIdrPass)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST__Dfct, &CMP15_Dfct_RcrdDlg::OnDblclkListDfct)
	ON_EN_KILLFOCUS(IDC_EDIT_Hide_Ipt, &CMP15_Dfct_RcrdDlg::OnKillfocusEditHideIpt)
	ON_BN_CLICKED(IDC_BUTTON_Start_Qry, &CMP15_Dfct_RcrdDlg::OnBnClickedButtonStartQry)
	ON_BN_CLICKED(IDC_BUTTON_Excel, &CMP15_Dfct_RcrdDlg::OnBnClickedButtonExcel)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST__Dfct, &CMP15_Dfct_RcrdDlg::OnCustomdrawListDfct)
END_MESSAGE_MAP()


// CMP15_Dfct_RcrdDlg 訊息處理常式

BOOL CMP15_Dfct_RcrdDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 設定此對話方塊的圖示。當應用程式的主視窗不是對話方塊時，
	// 框架會自動從事此作業
	SetIcon(m_hIcon, TRUE);			// 設定大圖示
	SetIcon(m_hIcon, FALSE);		// 設定小圖示

	// TODO: 在此加入額外的初始設定

	/*this->Db_Ctx.OnInitADOConn();
	Db_Ctx.GetRecordSet("select * from MP_Main_Tab where Mp_Mn_ID = 1 ;");
	_variant_t Var_Id = Db_Ctx.m_pRecordset->GetCollect("MP_Mn_ID");
	_variant_t Var_Dfct_Item_1 = Db_Ctx.m_pRecordset->GetCollect("MP_Mn_Dfct_Item_1");
	_variant_t Var_Input_Time = Db_Ctx.m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
	
	COleDateTime ip_tm = (COleDateTime)Var_Input_Time;
	int j = ip_tm.GetYear();
	int i = Var_Id;
	string Dfct_Item_1 = _bstr_t(Var_Dfct_Item_1);*/
	this->GetDlgItem(IDC_EDIT_Static_Time)->SetWindowText ( _T("至") );
	((CEdit*)this->GetDlgItem(IDC_EDIT_Static_Time))->SetReadOnly(true);

	Col_Name.push_back(_T("不良時間"));
	Col_Name.push_back(_T("箱號條碼"));
	Col_Name.push_back(_T("SN條碼"));
	Col_Name.push_back(_T("線別"));
	Col_Name.push_back(_T("不良現象1"));
	Col_Name.push_back(_T("不良現象2"));
	Col_Name.push_back(_T("Pass人"));
	Col_Name.push_back(_T("Pass時間"));
	for(std::vector<CString> ::reverse_iterator itr = Col_Name.rbegin();
		itr != Col_Name.rend();
		++itr)
	{
		m_CList.InsertColumn(0, (*itr) , LVCFMT_LEFT, 150);
	}

	/*m_CList.InsertColumn(0, _T("箱號條碼"), LVCFMT_LEFT, 190);
	m_CList.InsertColumn(1, _T("SN條碼"), LVCFMT_LEFT, 196);
	m_CList.InsertColumn(2, _T("線別"), LVCFMT_LEFT, 96);
	m_CList.InsertColumn(3, _T("不良現象1"), LVCFMT_LEFT, 196);
	m_CList.InsertColumn(4, _T("不良現象2"), LVCFMT_LEFT, 196);*/
	//m_CList.InsertColumn(5, _T("Pass人"), LVCFMT_LEFT, 196);
		//m_CList.InsertColumn(6, _T("Pass時間"), LVCFMT_LEFT, 196);
	//DWORD dwStyle =m_CList.GetExtendedStyle();

	

	//dwStyle |= LVS_EX_GRIDLINES;       //网格线（只适用与report风格）
	//dwStyle |= LVS_EX_GRIDLINES;
	m_CList.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	/*m_CList.InsertItem(LVIF_TEXT | LVIF_STATE, 0,
						_T("strResource"), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
					m_CList.SetItem(0, 1, LVIF_TEXT | LVIF_STATE,
						_T("strFirstPos"), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
					m_CList.SetItem(0, 2, LVIF_TEXT | LVIF_STATE,
						_T("strSecondPos"), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);*/



	// Hide the query field
	HCtrSw.push_back(this->GetDlgItem(IDC_RADIO_By_Time) );
	HCtrSw.push_back(this->GetDlgItem(IDC_DATETIMEPICKER_Str));	
	HCtrSw.push_back(this->GetDlgItem(IDC_EDIT_Static_Time) );
	HCtrSw.push_back(this->GetDlgItem(IDC_DATETIMEPICKER_End));
	HCtrSw.push_back(this->GetDlgItem(IDC_RADIO_By_Wrk_Numb) );
	HCtrSw.push_back(this->GetDlgItem(IDC_EDIT_Ord_Numb) );
	HCtrSw.push_back(this->GetDlgItem(IDC_BUTTON_Start_Qry) );
	//HCtrSw.push_back(this->GetDlgItem(IDC_BUTTON_ReQry) );
	for(ItrHCtrSw = HCtrSw.begin(); ItrHCtrSw != HCtrSw.end(); ++ ItrHCtrSw)
	{
		
		(*ItrHCtrSw)->ShowWindow(SW_HIDE);
	 	
	}



	((CComboBox*)this->GetDlgItem(IDC_COMBO_Line))->AddString(_T("A1"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Line))->AddString(_T("A2"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Line))->AddString(_T("A3"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Line))->AddString(_T("A4"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Line))->AddString(_T("線外"));
	//((CComboBox*)this->GetDlgItem(IDC_COMBO_Line))->DropDownStyle = ComboBoxStyle.DropDownList;

	//IDC_COMBO_Dfct_1
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("貼紙不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("包材不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("彩盒不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("外箱不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("欠板"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("欠料"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_1))->AddString(_T("其它"));

	//IDC_COMBO_Dfct_2
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("貼紙不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("包材不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("彩盒不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("外箱不良"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("欠板"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("欠料"));
	((CComboBox*)this->GetDlgItem(IDC_COMBO_Dfct_2))->AddString(_T("其它"));


	//Hide Pass_Edit 
	m_SubItemEdit.ShowWindow(SW_HIDE);


	//
	//select * from dbo.MP_Main_Tab where MP_Mn_Ps_Time != '1900-01-01 00:00:00.000';
	//select * from dbo.MP_Main_Tab where MP_Mn_Ps_Time = '1900-01-01 00:00:00.000';
	this->GetDfct();
	this->GetPss();
	this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
	this->m_Pss_Cnt = Lst_Data_Pass.size();
	this->m_Add_Cnt = m_Dfct_Cnt + m_Pss_Cnt;
	UpdateData(FALSE);
	 
	//
	this->Lst_Data_Dfct_Show_Lst();
	//this->m_Lst_State = _T("");
	//this->Lst_Data_Pss_Show_Lst();
	
	this->Get_Today_Dfct();
	this->Get_Today_Pass();
	
	return TRUE;  // 傳回 TRUE，除非您對控制項設定焦點
}

// 如果將最小化按鈕加入您的對話方塊，您需要下列的程式碼，
// 以便繪製圖示。對於使用文件/檢視模式的 MFC 應用程式，
// 框架會自動完成此作業。

void CMP15_Dfct_RcrdDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 繪製的裝置內容

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 將圖示置中於用戶端矩形
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 描繪圖示
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 當使用者拖曳最小化視窗時，
// 系統呼叫這個功能取得游標顯示。
HCURSOR CMP15_Dfct_RcrdDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CMP15_Dfct_RcrdDlg::OnRclickListDfct(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: 在此添加控件通知?理程序代?
	//CMenu* pPopMenu = menu.GetSubMenu(0);
	//pPopMenu->EnableMenuItem(ID_MENU_ITEM1     ,MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
	//pPopMenu->EnableMenuItem(ID_MENU_ITEM2     ,MF_BYCOMMAND | MF_ENABLED);
	NM_LISTVIEW *pNMListCtrl = (NM_LISTVIEW *)pNMHDR;
	int Cur_Row = pNMListCtrl->iItem;
	int Cur_Col = pNMListCtrl->iSubItem;
	CMenu obMenu;
	obMenu.LoadMenu(IDR_MENU1);

	CMenu* pPopupMenu = obMenu.GetSubMenu(0);
	ASSERT(pPopupMenu);

	// Get the cursor position
	CPoint obCursorPoint = (0, 0);

	GetCursorPos(&obCursorPoint);

	CString Pss_info= _T(""); 
	Pss_info = m_CList.GetItemText(this->Row_Edit_Idx, this->Col_Edit_Index);

	if (m_CList.GetSelectedCount() > 0 && _T("") != Pss_info && Cur_Row == Row_Edit_Idx )
	{
		pPopupMenu->EnableMenuItem(ID_IDR_PASS, MF_BYCOMMAND | MF_ENABLED);
		
	}
	else
	{
		pPopupMenu->EnableMenuItem(ID_IDR_PASS, MF_BYCOMMAND | MF_GRAYED | MF_DISABLED);
	}

	// Track the popup menu
	pPopupMenu->TrackPopupMenu(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON, obCursorPoint.x,
		obCursorPoint.y, this);

	*pResult = 0;
}


void CMP15_Dfct_RcrdDlg::OnBnClickedButtonQry()
{
	// TODO: 在此添加控件通知?理程序代?
	/*typedef std::vector<CTraceData> VctCTraceData;
typedef std::vector<CTraceData>::iterator VctCTraceDataIterator;*/
	/*if (6 > m_CList.GetHeaderCtrl()->GetItemCount())
	{
		 LONG styles;
          styles=GetWindowLong(m_CList.m_hWnd,GWL_STYLE);
          SetWindowLong(m_CList.m_hWnd,GWL_STYLE,styles | LVS_REPORT);
          
		m_CList.InsertColumn(5, _T("Pass人"), LVCFMT_LEFT, 196);
		m_CList.InsertColumn(6, _T("Pass時間"), LVCFMT_LEFT, 196); 
	}*/
	
	
	ItrHCtrSw = HCtrSw.begin();
	if((*ItrHCtrSw)->IsWindowVisible())
	{
		for(ItrHCtrSw = HCtrSw.begin(); ItrHCtrSw != HCtrSw.end(); ++ ItrHCtrSw)
		{
			if((*ItrHCtrSw)->IsWindowVisible())
			{
			(*ItrHCtrSw)->ShowWindow(SW_HIDE);
			}
			 	
		}
		this->Lst_Data_Dfct_Show_Lst();
		//this->m_Add_Cnt = this->Lst_Data_Dfct.size();
	}
	else /*if ( !(*ItrHCtrSw)->IsWindowVisible())*/
	{
		for(ItrHCtrSw = HCtrSw.begin(); ItrHCtrSw != HCtrSw.end(); ++ ItrHCtrSw)
		{
			if( !(*ItrHCtrSw)->IsWindowVisible() )
			{
			(*ItrHCtrSw)->ShowWindow(SW_SHOW);		
			}
		}
		this->Lst_Data_Pss_Show_Lst();

		/*this->m_Add_Cnt = this->Lst_Data_Pass.size();
		UpdateData(FALSE);*/
	}
}


void CMP15_Dfct_RcrdDlg::OnBnClickedButtonDft()
{
	// TODO: 在此添加控件通知?理程序代?
	//IDC_EDIT_Box_Numb
	//IDC_EDIT_SN_Numb
	//IDC_COMBO_Line
	//IDC_EDIT_Dfct_1
	//IDC_EDIT_Dfct_2
	//IDC_BUTTON_Cfm
	CString Focused_Edit_Text=_T("");
	CWnd* pFocused_Wnd = GetFocus();
	pFocused_Wnd->GetWindowTextW(Focused_Edit_Text);
	if( _T("")!= Focused_Edit_Text)
	{

		int Ctr_ID = pFocused_Wnd->GetDlgCtrlID();
		switch (Ctr_ID)
		{
			case (IDC_EDIT_Box_Numb):
				{
					UpdateData(TRUE);
					this->m_Box_Numb = Focused_Edit_Text;
					GetDlgItem(IDC_EDIT_SN_Numb)->SetFocus();
					break;
				}
			case (IDC_EDIT_SN_Numb):
				{
					UpdateData(TRUE);
					this->m_SN_Numb = Focused_Edit_Text;
					GetDlgItem(IDC_COMBO_Line)->SetFocus();
					break;
				}
			/*case (IDC_COMBO_Line):
				{
					this->m_Line = Focused_Edit_Text;
					GetDlgItem(IDC_EDIT_Dfct_1)->SetFocus();
					break;
				}*/
			default:
				break;
		}	// end of switch
	}	// end of if
}


void CMP15_Dfct_RcrdDlg::OnBnClickedButtonCfm()
{
	// TODO: 在此添加控件通知?理程序代?
	UpdateData(TRUE);

	CString Cur_Time = _T("");
	time_t t = time(0); 
	char tmp[64]; 
	strftime( tmp, sizeof(tmp), "%Y-%m-%d %H:%M:%S.000",localtime(&t) );  
	Cur_Time = tmp;

	 
	//CString Check_Box_Numb = _T("");
	//Check_Box_Numb = this->m_Box_Numb;
 
	
	// insert into MP_Main_Tab ( MP_Mn_Box_Numb,    MP_Mn_SSN_Numb,  MP_Mn_Ipt_Time,  MP_Mn_Ps_Time,  MP_Mn_Line,  
	//					MP_Mn_Dfct_Item_1,  MP_Mn_Dfct_Item_2, MP_Mn_Psser_Wrk_Numb, MP_Mn_Psser_Wrk_Name )
	//		values ('2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3',     '2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3' );
	if(   this->m_Box_Numb == _T("")
	    ||this->m_SN_Numb == _T("")
		||this->m_Line == _T("")
		||this->m_Dfct_1 == _T("")
		//||this->m_Dfct_2 == _T("")
		||( 1 != this->Check_Box_Numb() )
		)
	{
		::MessageBoxW(this->m_hWnd,_T("Waring: 未提交此項數據 , 此項數據有輸入项為空或箱号不符合指定规定! "),_T("Waring"),MB_OKCANCEL);
	}
	else{
		try{
			/*CString Base_Ins = _T( "insert into MP_Main_Tab ( MP_Mn_Box_Numb,    MP_Mn_SSN_Numb,  MP_Mn_Ipt_Time,  MP_Mn_Ps_Time, ")
						_T("MP_Mn_Line, MP_Mn_Dfct_Item_1,  MP_Mn_Dfct_Item_2, MP_Mn_Psser_Wrk_Numb, MP_Mn_Psser_Wrk_Name )")
						_T("values " ) ;
			CString Val = _T("");
			CString Snl = _T("'");
			Val = "門";
			this->Data_Base_Ctrl.ExecuteSQL((_bstr_t)(Base_Ins +  Snl+ Val +Snl )) ;
			
			this->Data_Base_Ctrl.ExecuteSQL(
						"insert into MP_Main_Tab ( MP_Mn_Box_Numb,    MP_Mn_SSN_Numb,  MP_Mn_Ipt_Time,  MP_Mn_Ps_Time, "
						"MP_Mn_Line, MP_Mn_Dfct_Item_1,  MP_Mn_Dfct_Item_2, MP_Mn_Psser_Wrk_Numb, MP_Mn_Psser_Wrk_Name )"
						"values " 
						"('2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3',  "
						" '2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3' );"
						); */
			

			Ctx_Db  Ctx_Db_Ctrl;
			Ctx_Db_Ctrl.MP_Mn_Box_Numb = this->m_Box_Numb;
			Ctx_Db_Ctrl.MP_Mn_SSN_Numb = this->m_SN_Numb;
			Ctx_Db_Ctrl.MP_Mn_Dfct_Item_1 = this->m_Dfct_1;
			Ctx_Db_Ctrl.MP_Mn_Dfct_Item_2 = this->m_Dfct_2;
			//Ctx_Db_Ctrl.MP_Mn_Extr = _T("2011-3-3");
			Ctx_Db_Ctrl.MP_Mn_Ipt_Time = Cur_Time;
			//Ctx_Db_Ctrl.MP_Mn_Ipt_Time = _T("2013-1-1");
			Ctx_Db_Ctrl.MP_Mn_Line = this->m_Line;
			//Ctx_Db_Ctrl.MP_Mn_Psser_Wrk_Name = _T("2011-3-3");
			//Ctx_Db_Ctrl.MP_Mn_Psser_Wrk_Numb = _T("2011-3-3");
			//Ctx_Db_Ctrl.MP_Mn_Ps_Time = _T("2011-3-3");
			//Ctx_Db_Ctrl.MP_Mn_Reserve_1 = _T("2011-3-3");
			//Ctx_Db_Ctrl.MP_Mn_Reserve_2 = _T("2011-3-3");
			//Ctx_Db_Ctrl.MP_Mn_Reserve_3 = _T("2011-3-3");
			Ctx_Db_Ctrl.Add_Item();
		}
		catch(...)
		{
			::MessageBoxW(this->m_hWnd,_T("Err: 遠端數據庫連接出錯, 數據并未提交, 請重新嘗試 , 可能是無網路連接,或者是遠端服務器出錯"),_T("Error"),MB_OKCANCEL);
		}

		List_Row_Data  Lst_rw_dt;

		Lst_rw_dt.MP_Mn_Box_Numb = this->m_Box_Numb;
		Lst_rw_dt.MP_Mn_SSN_Numb = this->m_SN_Numb;
		Lst_rw_dt.MP_Mn_Line = this->m_Line;
		Lst_rw_dt.MP_Mn_Dfct_Item_1 = this->m_Dfct_1;
		Lst_rw_dt.MP_Mn_Dfct_Item_2 = this->m_Dfct_2;
		Lst_rw_dt.MP_Mn_Ipt_Time = Cur_Time;
		if(0 != Lst_rw_dt.Push_Back_Lst_Ctrl(this->m_CList))
		{
			
			::MessageBoxW(this->m_hWnd,_T("Err: Fail in Adding ,Func Lst_rw_dt.Push_Back_Lst_Ctrl(CList&) fail"),_T("Error"),MB_OKCANCEL);
		}
		else{
			Lst_Data_Dfct.push_back(Lst_rw_dt);
			this->m_Dfct_Cnt = Lst_Data_Dfct.size();
			this->m_Pss_Cnt = this->Lst_Data_Pass.size();
			this->m_Add_Cnt = m_Dfct_Cnt + m_Pss_Cnt;

			this->m_Box_Numb = _T("");
			this->m_SN_Numb = _T("");
			this->m_Line = _T("");
			this->m_Dfct_1 = _T("");
			this->m_Dfct_2 = _T("");

			( (CComboBox*) GetDlgItem( IDC_COMBO_Dfct_1 ) )->SetCurSel( -1 );
			( (CComboBox*) GetDlgItem( IDC_COMBO_Line ) )->SetCurSel( -1 );

			UpdateData(FALSE);
		}

	}
	this->Get_Today_Dfct();
	this->Get_Today_Pass();
}




void CMP15_Dfct_RcrdDlg::OnBnClickedButtonClr()
{
	// TODO: 在此添加控件通知?理程序代?
	this->m_Box_Numb = _T("");
	this->m_SN_Numb = _T("");
	this->m_Line = _T("");
	this->m_Dfct_1 = _T("");
	this->m_Dfct_2 = _T("");
	
	GetDlgItem(IDC_EDIT_Box_Numb)->SetFocus();
	( (CComboBox*) GetDlgItem( IDC_COMBO_Dfct_1 ) )->SetCurSel( -1 );
	( (CComboBox*) GetDlgItem( IDC_COMBO_Line ) )->SetCurSel( -1 );
	UpdateData(FALSE);
}


// Menu_Pass
void CMP15_Dfct_RcrdDlg::OnIdrPass()
{

	// TODO: 在此添加命令?理程序代?
	POSITION SelectionPos = m_CList.GetFirstSelectedItemPosition();
	int iCurSel = -1;
	SelectionPos = m_CList.GetFirstSelectedItemPosition();
	iCurSel = m_CList.GetNextSelectedItem(SelectionPos);
	//List_Row_Data Lst_Data_Pass
	if( -1 != iCurSel )
	{
		
		if(iCurSel + 1 <= Lst_Data_Dfct.size())
		{
			
			CString Ipt_Time = Lst_Data_Dfct[iCurSel].MP_Mn_Ipt_Time;

			CString Cur_Pss_Time = _T("");
			time_t t = time(0); 
			char tmp[64]; 
			strftime( tmp, sizeof(tmp), "%Y-%m-%d %H:%M:%S.000",localtime(&t) );  
			Cur_Pss_Time = tmp;

			
			Ctx_Db  Ctx_Db_Ctrl;
			Ctx_Db_Ctrl.Update_Pass(Ipt_Time,this->m_Ipt_Pss_Info,Cur_Pss_Time);


			//m_Ipt_Pss_Info
			Lst_Data_Dfct[iCurSel].MP_Mn_Ps_Time = Cur_Pss_Time;
			Lst_Data_Dfct[iCurSel].MP_Mn_Psser_Wrk_Numb = this->m_Ipt_Pss_Info;
			List_Row_Data lrd = Lst_Data_Dfct[iCurSel];	 
			this->Lst_Data_Pass.push_back(Lst_Data_Dfct[iCurSel]);

			std::vector<List_Row_Data>::iterator itr;
			itr = Lst_Data_Dfct.begin()+iCurSel;	 
			this->Lst_Data_Dfct.erase(itr);


			this->Lst_Data_Dfct_Show_Lst();
			this->m_Pss_Cnt = this->Lst_Data_Pass.size();
			UpdateData(FALSE);
			//this->Lst_Data_Pss_Show_Lst();
		}
		else{
			::MessageBoxW(this->m_hWnd,_T("Err: Fail in Lst_Data_Dfct ,Index Parm is out of range"),_T("Error"),MB_OKCANCEL);
		}
	}
	else{
		::MessageBoxW(this->m_hWnd,_T("Err: Fail in Lst_Select ,not Select"),_T("Error"),MB_OKCANCEL);
	}

	this->Get_Today_Dfct();
	this->Get_Today_Pass();
}


void CMP15_Dfct_RcrdDlg::OnDblclkListDfct(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: 在此添加控件通知?理程序代?

	NM_LISTVIEW *pNMListCtrl = (NM_LISTVIEW *)pNMHDR;
	CString SetStr =_T("");
	SetStr = m_CList.GetItemText(pNMListCtrl->iItem, pNMListCtrl->iSubItem);
	int Row_Index = pNMListCtrl->iItem;
	int Col_Index = pNMListCtrl->iSubItem;
	if (pNMListCtrl->iItem != -1 && 6 == Col_Index   && 0 == Shw_Pss_Lst_Flg)
	{
		CRect rect, dlgRect;
		//鳳腕絞蹈腔遵僅眕巠茼彆蚚誧覃淕遵僅  
		//森揭祥蚚鳳腕腔赽砐rect撻倛遺懂扢离遵僅  
		//岆秪菴珨蹈奀殿隙腔遵僅岆淕俴腔遵僅  
		int width = this->m_CList.GetColumnWidth(pNMListCtrl->iSubItem);
		m_CList.GetSubItemRect(pNMListCtrl->iItem, pNMListCtrl->iSubItem, LVIR_BOUNDS, rect);
		//listSelFlagBase = pNMListCtrl->iItem;
		//listSelFlagSub = pNMListCtrl->iSubItem;
		m_CList.GetWindowRect(&dlgRect);
		ScreenToClient(&dlgRect);
		int height = rect.Height();
		rect.top += (dlgRect.top + 1);
		rect.left += (dlgRect.left + 1);
		rect.bottom = rect.top + height + 2;
		rect.right = rect.left + width + 2;

		m_SubItemEdit.MoveWindow(&rect);
		m_SubItemEdit.ShowWindow(SW_SHOW);
		m_SubItemEdit.SetFocus();
		this->Row_Edit_Idx = pNMListCtrl->iItem;
		this->Col_Edit_Index = pNMListCtrl->iSubItem;
		
		m_SubItemEdit.SetWindowText(SetStr);
	}

	*pResult = 0;
}


void CMP15_Dfct_RcrdDlg::OnKillfocusEditHideIpt()
{
	// TODO: 在此添加控件通知?理程序代?

	UpdateData(TRUE);
	if(_T("") != m_Ipt_Pss_Info)
	{
		m_CList.SetItem(Row_Edit_Idx, 6, LVIF_TEXT | LVIF_STATE,m_Ipt_Pss_Info, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		m_SubItemEdit.SetWindowText(_T(""));
		m_SubItemEdit.ShowWindow(SW_HIDE);
		//m_CList.DeleteItem(this->Row_Edit_Idx);
		//Lst_Data_Pass
	}
}


int CMP15_Dfct_RcrdDlg::GetDfct(void)
{
	this->Get_Today_Dfct();
	this->Get_Today_Pass();
	
	try{
		CString sql = _T("");  
		sql.Format(_T("select * from dbo.MP_Main_Tab where MP_Mn_Ps_Time = '1900-01-01 00:00:00.000' order by MP_Mn_Ipt_Time desc;"));  
		_RecordsetPtr m_pRecordset;  
		m_pRecordset =  this->Data_Base_Ctrl.GetRecordSet((_bstr_t)sql);
		while( 0 == this->Data_Base_Ctrl.m_pRecordset->adoEOF)
		{
		 
			List_Row_Data lst;
			lst.MP_Mn_Box_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Box_Numb");
			lst.MP_Mn_SSN_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_SSN_Numb");
			COleDateTime TemptimeIpt(m_pRecordset->GetCollect("MP_Mn_Ipt_Time"));
			CString szIpt =_T(""); 
			szIpt.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimeIpt.GetYear(),TemptimeIpt.GetMonth(), TemptimeIpt.GetDay()
													  ,TemptimeIpt.GetHour(),TemptimeIpt.GetMinute(),TemptimeIpt.GetSecond());
			lst.MP_Mn_Ipt_Time = szIpt;
			//lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			COleDateTime TemptimePss(m_pRecordset->GetCollect("MP_Mn_Ps_Time"));
			CString szPss =_T("");
			szPss.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimePss.GetYear(),TemptimePss.GetMonth(), TemptimePss.GetDay()
													  ,TemptimePss.GetHour(),TemptimePss.GetMinute(),TemptimePss.GetSecond());
			lst.MP_Mn_Ps_Time = szPss;
			//lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");
			lst.MP_Mn_Line= (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Line");

			lst.MP_Mn_Dfct_Item_1  = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_1");
			lst.MP_Mn_Dfct_Item_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_2");
			lst.MP_Mn_Psser_Wrk_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Numb");
			lst.MP_Mn_Psser_Wrk_Name = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Name");
			lst.MP_Mn_Extr = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Extr");

			lst.MP_Mn_Reserve_1 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_1");
			lst.MP_Mn_Reserve_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_2");
			lst.MP_Mn_Reserve_3 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_3");

			
			this->Lst_Data_Dfct.push_back(lst);
			m_pRecordset->MoveNext();  
		}
		this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
		UpdateData(FALSE);
		/*this->m_Pss_Cnt = Lst_Data_Pass.size();
		UpdateData(FALSE);*/
	}
	catch(...){
		::MessageBoxW(this->m_hWnd,_T("Err: Fail in GetDfct ,DateBase read fail"),_T("Error"),MB_OKCANCEL);
		return -1;
	}
	return 0;
}


int CMP15_Dfct_RcrdDlg::GetPss(void)
{
	this->Get_Today_Dfct();
	this->Get_Today_Pass();
	
	try{
		CString sql = _T("");  
		sql.Format(_T("select * from dbo.MP_Main_Tab where MP_Mn_Ps_Time != '1900-01-01 00:00:00.000' order by MP_Mn_Ipt_Time desc;"));  
		_RecordsetPtr m_pRecordset;  
		m_pRecordset =  this->Data_Base_Ctrl.GetRecordSet((_bstr_t)sql);
		while( 0 == this->Data_Base_Ctrl.m_pRecordset->adoEOF)
		{
		 
			List_Row_Data lst;
			lst.MP_Mn_Box_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Box_Numb");
			lst.MP_Mn_SSN_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_SSN_Numb");
			/*lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");*/
			COleDateTime TemptimeIpt(m_pRecordset->GetCollect("MP_Mn_Ipt_Time"));
			CString szIpt =_T(""); 
			szIpt.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimeIpt.GetYear(),TemptimeIpt.GetMonth(), TemptimeIpt.GetDay()
													  ,TemptimeIpt.GetHour(),TemptimeIpt.GetMinute(),TemptimeIpt.GetSecond());
			lst.MP_Mn_Ipt_Time = szIpt;
			//lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			COleDateTime TemptimePss(m_pRecordset->GetCollect("MP_Mn_Ps_Time"));
			CString szPss =_T("");
			szPss.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimePss.GetYear(),TemptimePss.GetMonth(), TemptimePss.GetDay()
													  ,TemptimePss.GetHour(),TemptimePss.GetMinute(),TemptimePss.GetSecond());
			lst.MP_Mn_Ps_Time = szPss;
			//lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");

			lst.MP_Mn_Line= (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Line");

			lst.MP_Mn_Dfct_Item_1  = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_1");
			lst.MP_Mn_Dfct_Item_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_2");
			lst.MP_Mn_Psser_Wrk_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Numb");
			lst.MP_Mn_Psser_Wrk_Name = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Name"); 
			lst.MP_Mn_Extr = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Extr");

			lst.MP_Mn_Reserve_1 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_1");
			lst.MP_Mn_Reserve_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_2");
			lst.MP_Mn_Reserve_3 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_3");


			this->Lst_Data_Pass.push_back(lst);
			m_pRecordset->MoveNext();  
		}
		/*this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
		UpdateData(FALSE);*/
		this->m_Pss_Cnt = Lst_Data_Pass.size();
		UpdateData(FALSE);
	}
	catch(...){
		::MessageBoxW(this->m_hWnd,_T("Err: Fail in GetPss ,DateBase read fail"),_T("Error"),MB_OKCANCEL);
		return -1;
	}
	return 0;
}


int CMP15_Dfct_RcrdDlg::Lst_Data_Dfct_Show_Lst(void)
{
	this->Get_Today_Dfct();
	this->Get_Today_Pass();
	
	Lst_Data_Dfct.clear();
	Lst_Data_Pass.clear();
	this->GetDfct();
	this->GetPss();
	this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
	this->m_Pss_Cnt = Lst_Data_Pass.size();
	this->m_Add_Cnt = m_Dfct_Cnt + m_Pss_Cnt;
	UpdateData(FALSE);


	std::vector<List_Row_Data>::iterator itr ;
	//Lst_Data_Dfct.clear();
	m_CList.DeleteAllItems();
	for( itr =  this->Lst_Data_Dfct.begin();  itr != Lst_Data_Dfct.end(); ++itr)
	{
		itr->Push_Back_Lst_Ctrl(m_CList);
	}
	this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;

	this->m_Lst_State = _T("未PASS明細");
	UpdateData(FALSE);
	Lst_Mry_Excel_Flg = 1;
	return 0;
}


int CMP15_Dfct_RcrdDlg::Lst_Data_Pss_Show_Lst(void)
{
	this->Get_Today_Dfct();
	this->Get_Today_Pass();
	
	Lst_Data_Dfct.clear();
	Lst_Data_Pass.clear();
	this->GetDfct();
	this->GetPss();
	this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
	this->m_Pss_Cnt = Lst_Data_Pass.size();
	this->m_Add_Cnt = m_Dfct_Cnt + m_Pss_Cnt;
	UpdateData(FALSE);


	std::vector<List_Row_Data>::iterator itr ;
	//Lst_Data_Pass.clear();
	m_CList.DeleteAllItems();
	for( itr =  this->Lst_Data_Pass.begin();  itr != Lst_Data_Pass.end(); ++itr)
	{
		itr->Push_Back_Lst_Ctrl(m_CList);
		//CString i = itr->MP_Mn_Ps_Time;
	}
	this->m_Pss_Cnt = Lst_Data_Pass.size();

	this->m_Lst_State = _T("PASS明細");
	UpdateData(FALSE);
	Lst_Mry_Excel_Flg = 2;
	return 0;
}


void CMP15_Dfct_RcrdDlg::OnBnClickedButtonStartQry()
{
	// TODO: 在此添加控件通知?理程序代?
	UpdateData(TRUE);
	//this->Lst_Data_Pss_Show_Lst();
	this->Lst_Data_Qry.clear();
	m_CList.DeleteAllItems();
	CString sql = _T("");  
	
	_RecordsetPtr m_pRecordset;  
	
	/*while( 0 == this->Data_Base_Ctrl.m_pRecordset->adoEOF)
	{
		m_pRecordset->MoveNext(); 
	}*/
	if(0 == this->m_Qry_Flg)
	{
		/*if( this->m_Start_Time == this->m_End_Time)
		{
				 ;
		}*/
		CString Str_Time =_T("");
		Str_Time.Format( _T("%.4d-%.2d-%.2d 00:00:00.000"),this->m_Start_Time.GetYear()
									,this->m_Start_Time.GetMonth(), this->m_Start_Time.GetDay() );
		CString End_Time =_T("");
		End_Time.Format( _T("%.4d-%.2d-%.2d 23:59:59.000"), this->m_End_Time.GetYear(),this->m_End_Time.GetMonth()
														, this->m_End_Time.GetDay()	);
		//是否為同一天
		if( this->m_Start_Time  ==  this->m_End_Time   &&  Str_Time ==  End_Time)
		{
			 COleDateTime tmp_dt = this->m_Start_Time;
			 COleDateTimeSpan dbdate(1,0,0,0);
			 tmp_dt = tmp_dt + dbdate;
			 End_Time.Format( _T("%.4d-%.2d-%.2d 00:00:00.000"),tmp_dt.GetYear()
									,tmp_dt.GetMonth(), tmp_dt.GetDay()   );
		}// end of if , none else

		CString Snl = _T("'");
		CString And = _T("and ");
		CString Low = _T(" MP_Mn_Ipt_Time <=");
		CString Line_End = _T(";");
		CString Base_Str = _T("select * from dbo.MP_Main_Tab where MP_Mn_Ipt_Time >=  ");  
		sql = Base_Str   + Snl + Str_Time   + Snl   + And + Low   +   Snl  + End_Time   + Snl  + Line_End;
	}
	else{
		//m__Ord_Numb
		//select * from dbo.MP_Main_Tab where SubString(MP_Mn_Box_Numb ,1,14) = 'UO2-G80048-001';
		CString Base_Str = _T("select * from dbo.MP_Main_Tab where SubString(MP_Mn_Box_Numb ,1,14) = '");
		CString End = _T("';");
		sql = Base_Str + m_Ord_Numb + End;
	}

	m_pRecordset =  this->Data_Base_Ctrl.GetRecordSet((_bstr_t)sql);
	while( 0 == this->Data_Base_Ctrl.m_pRecordset->adoEOF)
	{
		List_Row_Data lst;
			lst.MP_Mn_Box_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Box_Numb");
			lst.MP_Mn_SSN_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_SSN_Numb");
			/*lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");*/
			COleDateTime TemptimeIpt(m_pRecordset->GetCollect("MP_Mn_Ipt_Time"));
			CString szIpt =_T(""); 
			szIpt.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimeIpt.GetYear(),TemptimeIpt.GetMonth(), TemptimeIpt.GetDay()
													  ,TemptimeIpt.GetHour(),TemptimeIpt.GetMinute(),TemptimeIpt.GetSecond());
			lst.MP_Mn_Ipt_Time = szIpt;
			//lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			COleDateTime TemptimePss(m_pRecordset->GetCollect("MP_Mn_Ps_Time"));
			CString szPss =_T("");
			szPss.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimePss.GetYear(),TemptimePss.GetMonth(), TemptimePss.GetDay()
													  ,TemptimePss.GetHour(),TemptimePss.GetMinute(),TemptimePss.GetSecond());
			lst.MP_Mn_Ps_Time = szPss;
			//lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");

			lst.MP_Mn_Line= (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Line");

			lst.MP_Mn_Dfct_Item_1  = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_1");
			lst.MP_Mn_Dfct_Item_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_2");
			lst.MP_Mn_Psser_Wrk_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Numb");
			lst.MP_Mn_Psser_Wrk_Name = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Name"); 
			lst.MP_Mn_Extr = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Extr");

			lst.MP_Mn_Reserve_1 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_1");
			lst.MP_Mn_Reserve_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_2");
			lst.MP_Mn_Reserve_3 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_3");


			this->Lst_Data_Qry.push_back(lst);
		m_pRecordset->MoveNext(); 
	}
	int iPss = 0;
	int jDfct = 0;
	for( std::vector<List_Row_Data>::iterator itr = Lst_Data_Qry.begin();
		 itr != Lst_Data_Qry.end();
		 ++itr)
	{
		itr->Push_Back_Lst_Ctrl(this->m_CList);
		if(  _T("1900-01-01 00:00:00.000") != itr->MP_Mn_Ps_Time  )
		{
			iPss ++;
		}
		else{
			jDfct++;
		}

	}
	this->m_Dfct_Cnt = jDfct ;
	this->m_Pss_Cnt = iPss;
	this->m_Add_Cnt = Lst_Data_Qry.size();
	UpdateData(FALSE);
	Lst_Mry_Excel_Flg = 3;
}


void CMP15_Dfct_RcrdDlg::OnBnClickedButtonExcel()
{
	// TODO: 在此添加控件通知?理程序代?



			TCHAR path[255];
			SHGetSpecialFolderPathW(0, path,CSIDL_DESKTOPDIRECTORY,0);
			//MessageBox(path);
			CString sFileName =CString(path);
				CString tmp_Div_File = _T("\\");
				//sFileName = op.GetPathName();
				//sFileName = sFileName + tmp_Div_File + _T("Mp_15_Dfct_Lst_Exprt.xlsx");
				sFileName = sFileName + tmp_Div_File + _T("WUHAIBING.xls");

			CFileDialog op(TRUE, _T("xlsx"),
				sFileName,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				NULL, this);
			if (IDOK == op.DoModal())
			{
				//CString sFileName =_T("C:\\ProgramData\\DrawTrace\\TraceData.xlsx");
				//CString wstrPathName = op.GetPathName();
				CString wstrPathName = op.GetFolderPath();
				CString wstrFileName = op.GetFileName();
				CString wstrAbsName = wstrPathName + tmp_Div_File + wstrFileName;

				CString Dbg_Path_Name = op.GetPathName();
				CString Dbg_File_Name = op.GetFileName();
				if( ! ::PathFileExists(wstrAbsName) )
				{
					//::MessageBoxW(this->m_hWnd,_T("Wrn: 請選擇一個已經存在的EXCEL文件"),_T("Wrn"),MB_OKCANCEL);
					//::CreateFileW(sFileName,GENERIC_READ|GENERIC_WRITE,0,NULL,OPEN_ALWAYS,   FILE_FLAG_DELETE_ON_CLOSE,NULL); 
					CFile f( wstrAbsName, CFile::modeCreate | CFile::modeWrite ); 
					f.Close();
					
				}
				//CREATE_ALWAYS
				if( ! ::PathFileExists(wstrAbsName))  
				{
					/*CreateFileW(_T("sFileName"), GENERIC_READ || GENERIC_WRITE, FILE_SHARE_READ || FILE_SHARE_WRITE, NULL, CREATE_ALWAYS, NULL, NULL);
					CreateFileW(sFileName);  */
					::MessageBoxW(this->m_hWnd,_T("Wrn: 請選擇一個已經存在的EXCEL文件"),_T("Wrn"),MB_OKCANCEL);
					//return;
				}
				try
				{ 

					CoInitialize(NULL);

					Excel::_ApplicationPtr pExcelApp; 
					HRESULT hr = pExcelApp.CreateInstance(L"Excel.Application"); 
					ATLASSERT(SUCCEEDED(hr)); 
					pExcelApp->Visible = false;   // make Excel's main window visible 
					//Excel::_Workbook wrb = * (pExcelApp->Workbooks->Open("C:\\Users\\haibing.wu\\Desktop\\新增 Microsoft Excel 工作表.xlsx"));
					//Excel::_WorkbookPtr pWorkbook = pExcelApp->Workbooks->Open( (_bstr_t)sFileName);  // open excel file 
					Excel::_WorkbookPtr pWorkbook = pExcelApp->Workbooks->Open( (_bstr_t) (wstrAbsName)); 
					//C:\Users\haibing.wu\Desktop\Mp_15_Dfct_Lst_Exprt.xlsx
					Excel::_WorksheetPtr pWorksheet = pWorkbook->ActiveSheet; 
					pWorksheet->Name = _T("MP15"); 
					/*pWorksheet->Range [ "A1" ] [ vtMissing ]-> Value2 = _T("Exp A3");
					pWorksheet->Range [ "A2" ] [ vtMissing ]-> Value2 = _T("Exp A4");*/
					//pWorksheet->Cells->Item [ 1] [ 1 ]  = _T("Exp A9");
					CHeaderCtrl   *pmyHeaderCtrl= m_CList.GetHeaderCtrl(); //获取表头  

					std::vector<CString> Col_Name_Div;
					Col_Name_Div.push_back(_T("不良日期"));
					Col_Name_Div.push_back(_T("不良時間"));
					Col_Name_Div.push_back(_T("箱號條碼"));
					Col_Name_Div.push_back(_T("SN條碼"));
					Col_Name_Div.push_back(_T("線別"));
					Col_Name_Div.push_back(_T("不良現象1"));
					Col_Name_Div.push_back(_T("不良現象2"));
					Col_Name_Div.push_back(_T("Pass人"));
					Col_Name_Div.push_back(_T("Pass日期"));
					Col_Name_Div.push_back(_T("Pass時間"));

 
					int   m_cols   = pmyHeaderCtrl-> GetItemCount() + 2; //获取列数  
					int   m_rows = m_CList.GetItemCount();  //获取行数
					//Excel::Range Ptr_Range = pWorksheet->Cells;
					  for   (int   iCol   =   1;   iCol   <=   m_cols;   ++iCol)  //将列表内容写入EXCEL   
					  {   
						 
						   pWorksheet->Cells->PutColumnWidth(_variant_t((long)200/6));
						   pWorksheet->Cells->PutVerticalAlignment(_variant_t((short)-4108));

						  
						   pWorksheet->Cells->Item [ 1] [ iCol ] = (_variant_t)Col_Name_Div [iCol - 1];
						  
						   for( int  iRow   =   2;   iRow   <=   m_rows;   ++iRow)
						    
						   {   
								if( 1 == iCol || 9 == iCol)
								{
									if ( _T("") != m_CList.GetItemText(iRow-1,iCol-1))
									{
										if( 9 == iCol )
										{
											pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1-1).Left(10) );
										}
										else{
										pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1).Left(10) ); //取得m_CList控件中的内容  
										}
										//CString dbg = _T("");
										//dbg = m_CList.GetItemText(iRow-1,iCol-1).Left(10) ;
										//pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(dbg ); 
									}//if ( _T("") != m_CList.GetItemText(iRow-1,iCol-1))
									else{
										if( 9 == iCol )
										{
											pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1-1).Left(10) );
										}
										else{
										pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1).Left(10) );
										}
									}//if ( _T("") != m_CList.GetItemText(iRow-1,iCol-1))
								}
								
								else if( 2 == iCol || 10 == iCol)
								{
									if ( _T("") != m_CList.GetItemText(iRow-1,iCol-1))
									{
										if( 10 == iCol )
										{
											pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1 -1-1 ).Mid( 11, 5) );
										}
										else{
										pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1 -1 ).Mid( 11, 5) ); //取得m_CList控件中的内容  
										}
									}//if ( _T("") != m_CList.GetItemText(iRow-1,iCol-1))
									else{
										if( 10 == iCol )
										{
											pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1 -1 -1).Mid( 11, 5));
										}
										else{
										pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1 -1 ).Mid( 11, 5));
										}
									}//if ( _T("") != m_CList.GetItemText(iRow-1,iCol-1))
								}//else if( 2 == iCol || 11 == iCol)



								else if( 2 + 1 <= iCol && iCol <= 9 - 1 )
								{
									pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1-1));
								}

								else if( 10 + 1 <= iCol )
								{
									pWorksheet->Cells->Item [ iRow ] [ iCol ] = (_variant_t)(m_CList.GetItemText(iRow-1,iCol-1-1-1));
								}

						   }   
					  }    
					//Excel::RangePtr pRange = pWorksheet->Cells; 
 
					//const int nplot = 100; 
					//const double xlow = 0.0, xhigh = 20.0; 
					//double h = (xhigh-xlow)/(double)nplot; 
	
					//pRange -> Item[1][1] = L"x";  // read/write cell's data 
					//pRange -> Item[1][2] = L"f(x)"; 
					//for (int i=0;i<nplot;++i) 
					//{ 
					//    double x = xlow+i*h; 
					//    pRange->Item[i+2][1] = x; 
					//    pRange->Item[i+2][2] = sin(x)*exp(-x); 
					//} 

					//Excel::RangePtr pBeginRange = pRange->Item[1][1]; 
					//Excel::RangePtr pEndRange = pRange->Item[nplot+1][2]; 
					//Excel::RangePtr pTotalRange =  
					//    pWorksheet->Range[(Excel::Range*)pBeginRange][(Excel::Range*)pEndRange]; 
					//Excel::_ChartPtr pChart = pExcelApp->ActiveWorkbook->Charts->Add(); 
					//// refer to : 
					//// http://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.chart.chartwizard(v=vs.80).aspx 
					//pChart->ChartWizard( 
					//    (Excel::Range*)pTotalRange, 
					//    (long)Excel::xlXYScatter, 
					//    6L, 
					//    (long)Excel::xlColumns, 
					//    1L,1L, 
					//    true, 
					//    L"My Graph", 
					//    L"x",L"f(x)"); 
					//pChart->Name = L"My Data Plot"; 

					pWorkbook->Close(VARIANT_TRUE);  // save changes 
					pExcelApp->Quit(); 
				} 
				catch (_com_error& error) 
				{ 
					ATLASSERT(FALSE); 
					ATLTRACE2(error.ErrorMessage()); 
				}
			} //if (IDOK == op.DoModal())
				 
} // end of func
//
//				// CString sFileName = _T("E:\\myexcel.xls");  
//		 COleVariant covTrue((short)TRUE),covFalse((short)FALSE),covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);  
//		 
//		/* Excel::_ApplicationPtr pExcelApp; 
//		 Excel::_WorkbookPtr pWorkbook;
//		 Excel::_WorksheetPtr pWorksheet;
//		 Excel::RangePtr pRange ;
//*/
//		 CApplication0 app; //Excel程序  
//		 CWorkbooks0 books; //工作簿集合  
//		 CWorkbook0 book;  //工作表  
//		 CWorksheets0 sheets;  //工作簿集合  
//		 CWorksheet0 sheet; //工作表集合  
//		 CRange0 range; //使用区域  
//	
//		 //CoUninitialize();  
//		HRESULT hr; 
//		hr = CoInitialize(NULL); 
//		 //book.PrintPreview(_variant_t(false));  
//		 if(FAILED(hr))   
//		 {  
//			 ::MessageBoxW(this->m_hWnd,_T("Err: Failed to call Coinitialize()"),_T("Error"),MB_OKCANCEL);
//			  //MessageBox(_T("Failed to call Coinitialize()"));  
//			  return;  
//		 }  
//  
//		 if(!app.CreateDispatch(_T("Excel.Application.14"))); //创建IDispatch接口对象  
//		 {  
//			 ::MessageBoxW(this->m_hWnd,_T("Err: Failed to call CreateDispatch()"),_T("Error"),MB_OKCANCEL);
//			//MessageBox(_T("Error!"));  
//			 
//   
//		 }  
//		
//		//HRESULT hr = pExcelApp.CreateInstance(L"Excel.Application"); 
//		
//		//ATLASSERT(SUCCEEDED(hr)); 
//		 //"Excel.Application.14"
//		//if(!app.CreateDispatch(_T("Excel.Application.14"))); //创建IDispatch接口对象  
//		// {  
//		//    MessageBox(_T("Error!"));  
//  // 
//		// }  
//		 books = app.get_Workbooks();  
//		 book = books.Add(covOptional);  
//   
//  
//		 sheets = book.get_Worksheets();  
//		 sheet = sheets.get_Item(COleVariant((short)1));  //得到第一个工作表  
//  
//		 CHeaderCtrl   *pmyHeaderCtrl= m_CList.GetHeaderCtrl(); //获取表头  
//  
//			int   m_cols   = pmyHeaderCtrl-> GetItemCount(); //获取列数  
//			int   m_rows = m_CList.GetItemCount();  //获取行数  
//  
//  
//		 TCHAR     lpBuffer[256];      
//  
//		 HDITEM   hdi; //This structure contains information about an item in a header control. This structure has been updated to support header item images and order values.  
//			hdi.mask   =   HDI_TEXT;  
//			hdi.pszText   =   lpBuffer;  
//			hdi.cchTextMax   =   256;   
//  
//			CString   colname;  
//			CString strTemp;  
//  
//		 int   iRow,iCol;  
//		 for(iCol=0;   iCol <m_cols;   iCol++)//将列表的标题头写入EXCEL   
//		 {   
//  
//		  //GetCellName(1 ,iCol + 1, colname); //(colname就是对应表格的A1,B1,C1,D1)  
//			  colname = Col_Name[iCol];
//			  range   =   sheet.get_Range(COleVariant(colname),COleVariant(colname));    
//  
//			  pmyHeaderCtrl-> GetItem(iCol,   &hdi); //获取表头每列的信息  
//  
//			  range.put_Value2(COleVariant(hdi.pszText));  //设置每列的内容  
//  
//			  int   nWidth   =   m_CList.GetColumnWidth(iCol)/6;   
//  
//			  //得到第iCol+1列     
//			  range.AttachDispatch(range.get_Item(_variant_t((long)(iCol+1)),vtMissing).pdispVal,true);     
//  
//			  //设置列宽    
//			  range.put_ColumnWidth(_variant_t((long)nWidth));  
//  
//			}   
//  
//		  range   =   sheet.get_Range(COleVariant( _T("A1 ")),   COleVariant(colname));   
//  
//		  range.put_RowHeight(_variant_t((long)50));//设置行的高度   
//  
//  
//		  range.put_VerticalAlignment(COleVariant((short)-4108));//xlVAlignCenter   =   -4108   
//  
//		  COleSafeArray   saRet; //COleSafeArray类是用于处理任意类型和维数的数组的类  
//		  DWORD   numElements[]={m_rows,m_cols};       //行列写入数组  
//		  saRet.Create(VT_BSTR,   2,   numElements); //创建所需的数组  
//  
//		  range = sheet.get_Range(COleVariant( _T("A2 ")),covOptional); //从A2开始  
//		  range = range.get_Resize(COleVariant((short)m_rows),COleVariant((short)m_cols)); //表的区域  
//  
//		  long   index[2];    
//  
//		  for(   iRow   =   1;   iRow   <=   m_rows;   iRow++)//将列表内容写入EXCEL   
//		  {   
//		   for   (   iCol   =   1;   iCol   <=   m_cols;   iCol++)    
//		   {   
//			 index[0]=iRow-1;   
//			 index[1]=iCol-1;   
//  
//			 CString   szTemp;   
//  
//			 szTemp=m_CList.GetItemText(iRow-1,iCol-1); //取得m_CList控件中的内容  
//  
//			 BSTR   bstr   =   szTemp.AllocSysString(); //The AllocSysString method allocates a new BSTR string that is Automation compatible  
//  
//			 saRet.PutElement(index,bstr); //把m_CList控件中的内容放入saRet  
//  
//			 SysFreeString(bstr);  
//		   }   
//		  }    
//  
//      
//   
//			range.put_Value2(COleVariant(saRet)); //将得到的数据的saRet数组值放入表格  
//  
//  
//		 book.SaveCopyAs(COleVariant(sFileName)); //保存到sFileName文件  
//		 book.put_Saved(true);  
//  
//   
//  
//		 books.Close();  
//   
//		 //  
//   
//		 book.ReleaseDispatch();  
//		 books.ReleaseDispatch();   
//  
//		 app.ReleaseDispatch();  
//		 app.Quit(); 
//						 



int CMP15_Dfct_RcrdDlg::Get_Today_Dfct(void)
{
	/*CString Cur_Pss_Time = _T("");
			time_t t = time(0); 
			char tmp[64]; 
			strftime( tmp, sizeof(tmp), "%Y-%m-%d %00:%00:%00.000",localtime(&t) );  
			Cur_Pss_Time = tmp;*/
		/* COleDateTime Cur_dt;
		  COleDateTime Tomm_dt;
		Cur_dt = COleDateTime::GetCurrentTime();
		COleDateTimeSpan dbdate(1,0,0,0);
		Tomm_dt = Cur_dt + dbdate ;*/
	//SELECT * FROM dbo.MP_Main_Tab where MP_Mn_Ipt_Time >= '2016-08-29 00:00:00' and MP_Mn_Ipt_Time <= '2016-08-30 00:00:00';

	ToDay_Lst_Data_Dfct.clear();
	 COleDateTime Cur_dt;
		  COleDateTime Tomm_dt;
		Cur_dt = COleDateTime::GetCurrentTime();
		COleDateTimeSpan dbdate(1,0,0,0);
		Tomm_dt = Cur_dt + dbdate ;
		CString Cur_Day = _T("");
		CString Tom_Day = _T("");
		Cur_Day.Format(_T("%d-%d-%d 00:00:00.000"),Cur_dt.GetYear(),Cur_dt.GetMonth(),Cur_dt.GetDay());
		Tom_Day.Format(_T("%d-%d-%d 00:00:00.000"),Tomm_dt.GetYear(),Tomm_dt.GetMonth(),Tomm_dt.GetDay());
	try{
		CString sql = _T("");
		CString sql_Base_Str = _T("");  
		sql_Base_Str = (_T("select * from dbo.MP_Main_Tab where MP_Mn_Ps_Time = '1900-01-01 00:00:00.000' "));  
		CString Sql_Cnc_Tod = _T("");
		CString Sql_Cnc_Tmd = _T("");
		CString Sql_Cnc_Tmp_Cnt = _T("");
		CString Sql_Cnc_End_Cnt = _T("");
		Sql_Cnc_Tod = _T(" and MP_Mn_Ipt_Time >=  '");
		Sql_Cnc_Tmp_Cnt = _T("' and MP_Mn_Ipt_Time <=  '");
		Sql_Cnc_End_Cnt = _T("' order by MP_Mn_Ipt_Time desc;");
		sql = sql_Base_Str + Sql_Cnc_Tod + Cur_Day + Sql_Cnc_Tmp_Cnt + Tom_Day + Sql_Cnc_End_Cnt;
		_RecordsetPtr m_pRecordset;  
		m_pRecordset =  this->Data_Base_Ctrl.GetRecordSet((_bstr_t)sql);
		while( 0 == this->Data_Base_Ctrl.m_pRecordset->adoEOF)
		{
		 
			List_Row_Data lst;
			lst.MP_Mn_Box_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Box_Numb");
			lst.MP_Mn_SSN_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_SSN_Numb");
			/*lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");*/
			COleDateTime TemptimeIpt(m_pRecordset->GetCollect("MP_Mn_Ipt_Time"));
			CString szIpt =_T(""); 
			szIpt.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimeIpt.GetYear(),TemptimeIpt.GetMonth(), TemptimeIpt.GetDay()
													  ,TemptimeIpt.GetHour(),TemptimeIpt.GetMinute(),TemptimeIpt.GetSecond());
			lst.MP_Mn_Ipt_Time = szIpt;
			//lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			COleDateTime TemptimePss(m_pRecordset->GetCollect("MP_Mn_Ps_Time"));
			CString szPss =_T("");
			szPss.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimePss.GetYear(),TemptimePss.GetMonth(), TemptimePss.GetDay()
													  ,TemptimePss.GetHour(),TemptimePss.GetMinute(),TemptimePss.GetSecond());
			lst.MP_Mn_Ps_Time = szPss;
			//lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");

			lst.MP_Mn_Line= (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Line");

			lst.MP_Mn_Dfct_Item_1  = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_1");
			lst.MP_Mn_Dfct_Item_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_2");
			lst.MP_Mn_Psser_Wrk_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Numb");
			lst.MP_Mn_Psser_Wrk_Name = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Name"); 
			lst.MP_Mn_Extr = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Extr");

			lst.MP_Mn_Reserve_1 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_1");
			lst.MP_Mn_Reserve_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_2");
			lst.MP_Mn_Reserve_3 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_3");


			this->ToDay_Lst_Data_Dfct.push_back(lst);
			m_pRecordset->MoveNext();  
		}
		/*this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
		UpdateData(FALSE);*/
		this->m_Today_Dfct_Cnt = ToDay_Lst_Data_Dfct.size();
		UpdateData(FALSE);
	}
	catch(...){
		::MessageBoxW(this->m_hWnd,_T("Err: Fail in GetPss ,DateBase read fail"),_T("Error"),MB_OKCANCEL);
		return -1;
	}

	return 0;
}


int CMP15_Dfct_RcrdDlg::Get_Today_Pass(void)
{
	/*CString Cur_Pss_Time = _T("");
			time_t t = time(0); 
			char tmp[64]; 
			strftime( tmp, sizeof(tmp), "%Y-%m-%d %H:%M:%S.000",localtime(&t) );  
			Cur_Pss_Time = tmp;*/
	//SELECT * FROM dbo.MP_Main_Tab where MP_Mn_Ipt_Time >= '2016-08-29 00:00:00' and MP_Mn_Ipt_Time <= '2016-08-30 00:00:00';
	ToDay_Lst_Data_Pass.clear();
	 COleDateTime Cur_dt;
		  COleDateTime Tomm_dt;
		Cur_dt = COleDateTime::GetCurrentTime();
		COleDateTimeSpan dbdate(1,0,0,0);
		Tomm_dt = Cur_dt + dbdate ;
		CString Cur_Day = _T("");
		CString Tom_Day = _T("");
		Cur_Day.Format(_T("%d-%d-%d 00:00:00.000"),Cur_dt.GetYear(),Cur_dt.GetMonth(),Cur_dt.GetDay());
		Tom_Day.Format(_T("%d-%d-%d 00:00:00.000"),Tomm_dt.GetYear(),Tomm_dt.GetMonth(),Tomm_dt.GetDay());
	try{
		CString sql = _T("");
		CString sql_Base_Str = _T("");  
		sql_Base_Str = (_T("select * from dbo.MP_Main_Tab where MP_Mn_Ps_Time != '1900-01-01 00:00:00.000' "));  
		CString Sql_Cnc_Tod = _T("");
		CString Sql_Cnc_Tmd = _T("");
		CString Sql_Cnc_Tmp_Cnt = _T("");
		CString Sql_Cnc_End_Cnt = _T("");
		Sql_Cnc_Tod = _T(" and MP_Mn_Ipt_Time >= '");
		Sql_Cnc_Tmp_Cnt = _T("' and MP_Mn_Ipt_Time <= '");
		Sql_Cnc_End_Cnt = _T("' order by MP_Mn_Ipt_Time desc;");
		sql = sql_Base_Str + Sql_Cnc_Tod + Cur_Day + Sql_Cnc_Tmp_Cnt + Tom_Day + Sql_Cnc_End_Cnt;
		_RecordsetPtr m_pRecordset;  
		m_pRecordset =  this->Data_Base_Ctrl.GetRecordSet((_bstr_t)sql);
		while( 0 == this->Data_Base_Ctrl.m_pRecordset->adoEOF)
		{
		 
			List_Row_Data lst;
			lst.MP_Mn_Box_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Box_Numb");
			lst.MP_Mn_SSN_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_SSN_Numb");
			/*lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");*/
			COleDateTime TemptimeIpt(m_pRecordset->GetCollect("MP_Mn_Ipt_Time"));
			CString szIpt =_T(""); 
			szIpt.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimeIpt.GetYear(),TemptimeIpt.GetMonth(), TemptimeIpt.GetDay()
													  ,TemptimeIpt.GetHour(),TemptimeIpt.GetMinute(),TemptimeIpt.GetSecond());
			lst.MP_Mn_Ipt_Time = szIpt;
			//lst.MP_Mn_Ipt_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ipt_Time");
			COleDateTime TemptimePss(m_pRecordset->GetCollect("MP_Mn_Ps_Time"));
			CString szPss =_T("");
			szPss.Format( _T("%.4d-%.2d-%.2d %.2d:%.2d:%.2d.000"), TemptimePss.GetYear(),TemptimePss.GetMonth(), TemptimePss.GetDay()
													  ,TemptimePss.GetHour(),TemptimePss.GetMinute(),TemptimePss.GetSecond());
			lst.MP_Mn_Ps_Time = szPss;
			//lst.MP_Mn_Ps_Time = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Ps_Time");

			lst.MP_Mn_Line= (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Line");

			lst.MP_Mn_Dfct_Item_1  = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_1");
			lst.MP_Mn_Dfct_Item_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Dfct_Item_2");
			lst.MP_Mn_Psser_Wrk_Numb = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Numb");
			lst.MP_Mn_Psser_Wrk_Name = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Psser_Wrk_Name"); 
			lst.MP_Mn_Extr = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Extr");

			lst.MP_Mn_Reserve_1 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_1");
			lst.MP_Mn_Reserve_2 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_2");
			lst.MP_Mn_Reserve_3 = (TCHAR*)(_bstr_t)m_pRecordset->GetCollect("MP_Mn_Reserve_3");


			this->ToDay_Lst_Data_Pass.push_back(lst);
			m_pRecordset->MoveNext();  
		}
		/*this->m_Dfct_Cnt = Lst_Data_Dfct.size() ;
		UpdateData(FALSE);*/
		this->m__Today_Pass_Cnt = ToDay_Lst_Data_Pass.size();
		UpdateData(FALSE);
	}
	catch(...){
		::MessageBoxW(this->m_hWnd,_T("Err: Fail in GetPss ,DateBase read fail"),_T("Error"),MB_OKCANCEL);
		return -1;
	}
	return 0;
}


void CMP15_Dfct_RcrdDlg::OnCustomdrawListDfct(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: 在此添加控件通知?理程序代?
	LPNMLVCUSTOMDRAW lpLVCustomDraw = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);
	switch(lpLVCustomDraw->nmcd.dwDrawStage)
  {
    case CDDS_ITEMPREPAINT:
    case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
      if (0 == ((lpLVCustomDraw->nmcd.dwItemSpec + lpLVCustomDraw->iSubItem) % 2))
      {
        //lpLVCustomDraw->clrText = RGB(255,255,255); // white text
        //lpLVCustomDraw->clrTextBk = RGB(0,0,0); // black background

		lpLVCustomDraw->clrText = RGB(0,0,0); // white text
        lpLVCustomDraw->clrTextBk = RGB(240,240,240); // black background
      }
      else 
      {
        /*lpLVCustomDraw->clrText = CLR_DEFAULT;
        lpLVCustomDraw->clrTextBk = CLR_DEFAULT;*/

		lpLVCustomDraw->clrText = RGB(0,0,0); // white text
        lpLVCustomDraw->clrTextBk = RGB(255,255,255); // black background
      }
    break;

    default: break;    
  }


	*pResult = 0;

	*pResult |= CDRF_NOTIFYPOSTPAINT;
    *pResult |= CDRF_NOTIFYITEMDRAW;
    *pResult |= CDRF_NOTIFYSUBITEMDRAW;
}


int CMP15_Dfct_RcrdDlg::Check_Box_Numb(void)
{
	int iRtn = 0;
	if ( this->m_Box_Numb == _T("") )
	{
		::MessageBoxW(this->m_hWnd,_T("Err: 箱号为空"),_T("Error"),MB_OKCANCEL);
		iRtn = -1;
	}
	else if ( -1 == this->m_Box_Numb.Find('$') )
	{
		::MessageBoxW(this->m_hWnd,_T("Err: 箱号无$符"),_T("Error"),MB_OKCANCEL);
		iRtn =  -2;
	}
	else if ( 14 != this->m_Box_Numb.Left( m_Box_Numb.Find('$') ).GetLength() )
	{
		::MessageBoxW(this->m_hWnd,_T("Err: 箱号$符前字符串（订单号）不足十四位（订单号）"),_T("Error"),MB_OKCANCEL);
		iRtn =  -2;
	}

	else if ( 25 != this->m_Box_Numb.Left( m_Box_Numb.ReverseFind('$') ).GetLength() )
	{
		::MessageBoxW(this->m_hWnd,_T("Err: 箱号不足二十五位"),_T("Error"),MB_OKCANCEL);
		iRtn =  -3;
	}

	else
	{
		iRtn = 1;
	}
	return iRtn;
}
