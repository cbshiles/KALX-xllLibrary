// macro.cpp - Rename this file and replace this description.
// You need this: http://support.microsoft.com/kb/128185
#include "header.h"

using namespace xll;
                        
static AddInX xai_macro(
	// C function name, Excel macro name
	_T("?xll_macro"), _T("XLL.MACRO")
);
int WINAPI
xll_macro(void)
{
#pragma XLLEXPORT
	try {
		// macro body goes here
		ExcelX(xlcAlert, OPERX(_T("Hello World!")));
	}
	catch (const std::exception& ex) {
		// Use the macro ALERT.FILTER to control which alerts you see.
		XLL_ERROR(ex.what()); // or XLL_WARNING or XLL_INFO

		return 0; // failure
	}

	return 1; // success
}	