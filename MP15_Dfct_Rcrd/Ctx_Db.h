#pragma once
#include "stdafx.h"
//#include "mp15_dfct_rcrddlg.h"

#include "ADO_Conn.h"
//#include "MP15_Dfct_Rcrd.rc"
//using namespace std;
#include "vector"
typedef std::vector<CString> Vct_MP_Main_Tab;

class Ctx_Db :
//	public CMP15_Dfct_RcrdDlg,
	public ADOConn
{
public:
	Ctx_Db(void);
	~Ctx_Db(void);
	

public:
	CString MP_Mn_Box_Numb;
	CString MP_Mn_SSN_Numb;
	CString MP_Mn_Ipt_Time;
	CString MP_Mn_Ps_Time;
	CString MP_Mn_Line;

	CString MP_Mn_Dfct_Item_1;
	CString MP_Mn_Dfct_Item_2;
	CString MP_Mn_Psser_Wrk_Numb;
	CString MP_Mn_Psser_Wrk_Name;
	CString MP_Mn_Extr;

	CString MP_Mn_Reserve_1;
	CString MP_Mn_Reserve_2;
	CString MP_Mn_Reserve_3;

private:
	CString Div  ;
	CString Snl   ;
	CString Pts_L  ;
	CString Pts_R  ;
	CString Cmd_End  ;
public:
	//Vct_MP_Main_Tab  DtBs_MP_Main_Tab;
	//CMP15_Dfct_RcrdDlg
	//typedef std::vector<CTraceData> VctCTraceData;
	//typedef std::vector<CString> Vct_MP_Main_Tab;
	/* char szCurrentDateTime[32];     
time_t nowtime;     
struct tm* ptm;     
time(&nowtime);     
ptm = localtime(&nowtime);     
sprintf(szCurrentDateTime, "%4d-%.2d-%.2d %.2d:%.2d:%.2d",     
    ptm->tm_year + 1900, ptm->tm_mon + 1, ptm->tm_mday,     
    ptm->tm_hour, ptm->tm_min, ptm->tm_sec);*/
	Ctx_Db(Vct_MP_Main_Tab  Data_Base_Data); //throw int, data counts do not match
	int Delete_Item(CString Ipu_Time);
	int Delete_Item(int ID);
	int Add_Item();
	int Update_Item(Vct_MP_Main_Tab Data_Base_Data);
	int Update_Pass(CString Ipt_Time, CString Pss_Wrk_Numb, CString Pss_Time , CString Pss_Wrk_Name = _T(""));
};

