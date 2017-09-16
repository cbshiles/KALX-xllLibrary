// macro.cpp - Rename this file and replace this description.
// Uncomment the following line to use features for Excel2007 and above.
//#define EXCEL12
#include "xll/xll.h"

using namespace xll;

static AddInX xai_macro(_T("?xll_macro"), _T("XLL.MACRO"));
int WINAPI
xll_macro(void)
{
#pragma XLLEXPORT
	// macro body goes here

	return 1; // or 0 on failure
}	