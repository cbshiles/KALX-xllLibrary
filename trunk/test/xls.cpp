// xls.cpp - create spreadsheet programmatically
//#define EXCEL12
#include "xll/xll.h"

using namespace xll;

static AddInX xai_xls(_T("?xll_xls"), _T("XLL.XLS"));
int WINAPI
xll_xls(void)
{
#pragma XLLEXPORT
	
	try {
		XLL_XLC(New, OPERX(1)); // new 1 sheet workbook
		XLL_XLC(SaveAs, OPERX(_T("NewSheet.xls")));

		DefineStyleFont(_T("MyStyle"), _T("Courier"), 8);
		DefineStyleBorder(_T("MyStyle"), BT_OUTLINE);

		XLL_XLC(Select, XLL_XLF(Textref, OPERX(_T("D4")), OPERX(true)));
		Formula(_T("Cell D4"));
//		XLL_XLC(Formula, OPERX(_T("Cell D4")));
		XLL_XLC(ApplyStyle, OPERX(_T("MyStyle")));


		XLL_MOVE(1, 0);
		XLL_XLC(Formula, OPERX(_T("=R[-1]C[0]&\" is above this cell\"")));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
