// function.cpp - Rename this file and replace this description.
#include "header.h"

using namespace xll;

static AddInX xai_function(
	FunctionX(XLL_LPOPERX, _T("?xll_function"), _T("XLL.FUNCTION"))
	.Arg(XLL_DOUBLEX, _T("Number"), _T("is a number"))
//	.Arg(...) // add more args here
	.Category(CATEGORY)
	.FunctionHelp(_T("Description of what function does."))
	.Documentation(_T("Optional documentation for function."))
);
LPOPERX WINAPI
xll_function(double x)
{
#pragma XLLEXPORT
	static OPERX xResult;

	// function body goes here
	xResult = x; // Wait, what? Yes, an OPER can act like a double!
	
	return &xResult;
}