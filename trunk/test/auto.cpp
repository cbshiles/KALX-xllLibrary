// auto.cpp test Auto classes
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"

using namespace xll;

static AddInX xai_auto_example(
	_T("?xll_auto_example"), XLL_LPXLOPERX XLL_WORDX XLL_WORDX,
	_T("AUTO.EXAMPLE"), _T("rows, columns")
);
LPOPERX WINAPI
xll_auto_example(WORD r, WORD c)
{
#pragma XLLEXPORT
	OPERX* px = new OPERX(r,c);

//	x.resize(r, c);


	return px->dllfree();
}

static AddInX xai_xlfree_example(
	_T("?xll_xlfree_example"), XLL_LPXLOPERX XLL_WORDX XLL_WORDX,
	_T("XLFREE.EXAMPLE"), _T("")
);
LPXLOPERX WINAPI
xll_xlfree_example(WORD r, WORD c)
{
#pragma XLLEXPORT
	static LOPERX o;

	o = ExcelX(xlUDF, OPERX(_T("AUTO.EXAMPLE")), OPERX(r), OPERX(c));

	// same as o.xltype |= xlbitXLFree; return &o;
	return o.xlfree();
}
