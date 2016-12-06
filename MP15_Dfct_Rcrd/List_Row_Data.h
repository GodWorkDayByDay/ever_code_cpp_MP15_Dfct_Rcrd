#pragma once
#include "stdafx.h"
//#include "Windowss.h"

class List_Row_Data
{
public:
	List_Row_Data(void);
	virtual ~List_Row_Data(void);

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

public:
	 
	//List_Row_Data(CString Box_Numb, CString SN_Numb, CString Dfct_1, CString Dfct_2, CString Line, CString Ipt_Time);
	int Push_Back_Lst_Ctrl(CListCtrl& Lst_Ctrl);
};

