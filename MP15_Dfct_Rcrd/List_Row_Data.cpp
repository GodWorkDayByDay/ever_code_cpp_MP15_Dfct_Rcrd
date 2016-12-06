#include "StdAfx.h"
#include "List_Row_Data.h"


List_Row_Data::List_Row_Data(void)
{
		this-> MP_Mn_Box_Numb = _T("");
		this-> MP_Mn_SSN_Numb = _T("");
		this-> MP_Mn_Ipt_Time = _T("");
		this-> MP_Mn_Ps_Time = _T("");
		this-> MP_Mn_Line = _T("");

		this-> MP_Mn_Dfct_Item_1 = _T("");
		this-> MP_Mn_Dfct_Item_2 = _T("");
		this-> MP_Mn_Psser_Wrk_Numb = _T("");
		this-> MP_Mn_Psser_Wrk_Name = _T("");
		this-> MP_Mn_Extr = _T("");

		this-> MP_Mn_Reserve_1 = _T("");
		this-> MP_Mn_Reserve_2 = _T("");
		this-> MP_Mn_Reserve_3 = _T("");
	
}


List_Row_Data::~List_Row_Data(void)
{
}


//List_Row_Data::List_Row_Data(CString Box_Numb, CString SN_Numb, CString Dfct_1, CString Dfct_2, CString Line,CString Ipt_Time)
//{
//	this-> MP_Mn_Box_Numb = _T("");
//	this-> MP_Mn_SSN_Numb = _T("");
//	this-> MP_Mn_Ipt_Time = _T("");
//	this-> MP_Mn_Ps_Time = _T("");
//	this-> MP_Mn_Line = _T("");
//
//	this-> MP_Mn_Dfct_Item_1 = _T("");
//	this-> MP_Mn_Dfct_Item_2 = _T("");
//	this-> MP_Mn_Psser_Wrk_Numb = _T("");
//	this-> MP_Mn_Psser_Wrk_Name = _T("");
//	this-> MP_Mn_Extr = _T("");
//
//	this-> MP_Mn_Reserve_1 = _T("");
//	this-> MP_Mn_Reserve_2 = _T("");
//	this-> MP_Mn_Reserve_3 = _T("");
//
//
//
//	this-> MP_Mn_Box_Numb = Box_Numb;
//	this-> MP_Mn_SSN_Numb = SN_Numb;
//	this-> MP_Mn_Line = Line;
//	this-> MP_Mn_Dfct_Item_1 = Dfct_1;
//	this-> MP_Mn_Dfct_Item_2 = Dfct_2;
//	this-> MP_Mn_Ipt_Time = Ipt_Time;
//}


int List_Row_Data::Push_Back_Lst_Ctrl(CListCtrl& Lst_Ctrl)
{
	try{
		int iCount = Lst_Ctrl.GetItemCount();
		Lst_Ctrl.InsertItem(LVIF_TEXT | LVIF_STATE, iCount,this->MP_Mn_Ipt_Time, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		Lst_Ctrl.SetItem(iCount, 1, LVIF_TEXT | LVIF_STATE, this->MP_Mn_Box_Numb, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		Lst_Ctrl.SetItem(iCount, 2, LVIF_TEXT | LVIF_STATE,this->MP_Mn_SSN_Numb, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		Lst_Ctrl.SetItem(iCount, 3, LVIF_TEXT | LVIF_STATE,this->MP_Mn_Line, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		Lst_Ctrl.SetItem(iCount, 4, LVIF_TEXT | LVIF_STATE,this->MP_Mn_Dfct_Item_1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		Lst_Ctrl.SetItem(iCount, 5, LVIF_TEXT | LVIF_STATE,this->MP_Mn_Dfct_Item_2, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		if( _T("1900-01-01 00:00:00.000") != this->MP_Mn_Ps_Time  )
		{
		Lst_Ctrl.SetItem(iCount, 6, LVIF_TEXT | LVIF_STATE,this->MP_Mn_Psser_Wrk_Numb, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		Lst_Ctrl.SetItem(iCount, 7, LVIF_TEXT | LVIF_STATE,this->MP_Mn_Ps_Time, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED, 0, 0);
		}
		else{
			//null
		}
	}
	catch(...){
		return -1;
	}
	return 0;
}
