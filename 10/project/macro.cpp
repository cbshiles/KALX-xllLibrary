// macro.cpp - Rename this file and replace this description.
// You need this: http://support.microsoft.com/kb/128185
#include "header.h"

using namespace xll;

static AddInX xai_macro(_T("?xll_macro"), _T("XLL.MACRO"));
int WINAPI
xll_macro(void)
{
#pragma XLLEXPORT
	// macro body goes here
	ExcelX(xlcAlert, OPERX(_T("Hello World!")));

	return 1; // or 0 on failure
}	