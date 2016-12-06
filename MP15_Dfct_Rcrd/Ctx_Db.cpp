#include "StdAfx.h"
#include "Ctx_Db.h"


Ctx_Db::Ctx_Db(void)
{
	this->Div = _T(",");
	this->Snl = _T("'");
	this->Pts_L = _T(" ( ");
	this->Pts_R = _T(" ) ");
	this->Cmd_End = _T(" ; ");

	 
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


Ctx_Db::~Ctx_Db(void)
{
	this->ExitConnect();
}


Ctx_Db::Ctx_Db(Vct_MP_Main_Tab Data_Base_Data)
{
	if(Data_Base_Data.size() > 13)
	{
		throw 13;
	}
	else{
		//null value
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



		// value
		this-> MP_Mn_Box_Numb = Data_Base_Data[0];
		this-> MP_Mn_SSN_Numb = Data_Base_Data[1];
		this-> MP_Mn_Ipt_Time = Data_Base_Data[2];
		this-> MP_Mn_Ps_Time = Data_Base_Data[3];
		this-> MP_Mn_Line = Data_Base_Data[4];

		this-> MP_Mn_Dfct_Item_1 = Data_Base_Data[5];
		this-> MP_Mn_Dfct_Item_2 = Data_Base_Data[6];
		this-> MP_Mn_Psser_Wrk_Numb = Data_Base_Data[7];
		this-> MP_Mn_Psser_Wrk_Name = Data_Base_Data[8];
		this-> MP_Mn_Extr = Data_Base_Data[9];

		this-> MP_Mn_Reserve_1 = Data_Base_Data[10];
		this-> MP_Mn_Reserve_2 = Data_Base_Data[11];
		this-> MP_Mn_Reserve_3 = Data_Base_Data[12];

	} //end of else

}


int Ctx_Db::Delete_Item(CString Ipu_Time)
{

	return 0;
}


int Ctx_Db::Delete_Item(int ID)
{
	return 0;
}


int Ctx_Db::Add_Item()
{
	/*this->Data_Base_Ctrl.ExecuteSQL(
						"insert into MP_Main_Tab ( MP_Mn_Box_Numb,    MP_Mn_SSN_Numb,  MP_Mn_Ipt_Time,  MP_Mn_Ps_Time, "
						"MP_Mn_Line, MP_Mn_Dfct_Item_1,  MP_Mn_Dfct_Item_2, MP_Mn_Psser_Wrk_Numb, MP_Mn_Psser_Wrk_Name )"
						"values " 
						"('2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3',  "
						" '2011-3-3', '2011-3-3', '2011-3-3', '2011-3-3' );"
						); */

	CString Base_Ins = _T( "insert into MP_Main_Tab ( MP_Mn_Box_Numb,    MP_Mn_SSN_Numb,  MP_Mn_Ipt_Time,  MP_Mn_Ps_Time, ")
						_T("MP_Mn_Line, MP_Mn_Dfct_Item_1,  MP_Mn_Dfct_Item_2, MP_Mn_Psser_Wrk_Numb, MP_Mn_Psser_Wrk_Name, ")
						_T( "MP_Mn_Extr, MP_Mn_Reserve_1, MP_Mn_Reserve_2, MP_Mn_Reserve_3 )" )
						_T( " values " ) ;
			/*CString Div = _T(",");
			CString Snl = _T("'");
			CString Pts_L = _T(" ( ");
			CString Pts_R = _T(" ) ");
			CString Cmd_End = _T(" ; ");*/
			this->ExecuteSQL(  (_bstr_t) (Base_Ins +  Pts_L 
												   +  Snl+ this->MP_Mn_Box_Numb +Snl + Div
												   +  Snl+ this->MP_Mn_SSN_Numb +Snl + Div
												   +  Snl+ this->MP_Mn_Ipt_Time +Snl + Div
												   +  Snl+ this->MP_Mn_Ps_Time +Snl + Div
												   +  Snl+ this->MP_Mn_Line +Snl + Div
												   
												   +  Snl+ this->MP_Mn_Dfct_Item_1 + Snl + Div
												   +  Snl+ this->MP_Mn_Dfct_Item_2 + Snl + Div
												   +  Snl+ this->MP_Mn_Psser_Wrk_Numb + Snl + Div
												   +  Snl+ this->MP_Mn_Psser_Wrk_Name + Snl + Div
												   +  Snl+ this->MP_Mn_Extr + Snl + Div

												   +  Snl+ this->MP_Mn_Reserve_1 + Snl + Div
												   +  Snl+ this->MP_Mn_Reserve_2 + Snl + Div
												   +  Snl+ this->MP_Mn_Reserve_3 + Snl
												   +  Pts_R
												   +  Cmd_End
												   ) );

	return 0;
}


int Ctx_Db::Update_Item(Vct_MP_Main_Tab Data_Base_Data)
{
	return 0;
}


int Ctx_Db::Update_Pass(CString Ipt_Time, CString Pss_Wrk_Numb, CString Pss_Time ,CString Pss_Wrk_Name)
{
	//UPDATE Person SET FirstName = 'Fred' WHERE LastName = 'Wilson'
	CString Base_Upt;
	Base_Upt = _T("");
	Base_Upt = _T("update  MP_Main_Tab set ");


	CString	MP_Mn_Ps_Time_Equ = _T("");
	MP_Mn_Ps_Time_Equ = _T("MP_Mn_Ps_Time = ");

	CString	MP_Mn_Psser_Wrk_Numb_Equ = _T("");
	MP_Mn_Psser_Wrk_Numb_Equ = _T("MP_Mn_Psser_Wrk_Numb = ");

	CString	MP_Mn_Psser_Wrk_Name_Equ = _T("");
	MP_Mn_Psser_Wrk_Name_Equ = _T("MP_Mn_Psser_Wrk_Name = ");
	 
	CString Whr_Ipt_Time_Equ = _T("");
	Whr_Ipt_Time_Equ = _T("where MP_Mn_Ipt_Time = ");
	
	this->ExecuteSQL(  (_bstr_t) (Base_Upt +  MP_Mn_Ps_Time_Equ 
										   +  Snl+ Pss_Time +Snl + Div

										   +  MP_Mn_Psser_Wrk_Numb_Equ
										   +  Snl+ Pss_Wrk_Numb +Snl + Div

										   +  MP_Mn_Psser_Wrk_Name_Equ
										   +  Snl+ Pss_Wrk_Name +Snl
												

										   +  Whr_Ipt_Time_Equ
										   +  Snl+ Ipt_Time + Snl
										  
										   +  Cmd_End
												   ) );

	return 0;
}
