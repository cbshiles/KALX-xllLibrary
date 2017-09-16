// auto.cpp test Auto classes
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"

using namespace xll;

void
auto_free(void* px1)
{
	LPOPERX px = (LPOPERX)px1;

	ensure ((px->xltype&~xlbitDLLFree) == xltypeMulti);

	delete px->val.array.lparray;

	px->xltype = xltypeMissing; // like xlFree
}
//static Auto<FreeX> xaf_auto_example(auto_free);

static AddInX xai_auto_example(
	_T("?xll_auto_example"), XLL_LPXLOPERX XLL_WORDX XLL_WORDX,
	_T("AUTO.EXAMPLE"), _T("rows, columns")
);
LPXLOPERX WINAPI
xll_auto_example(WORD r, WORD c)
{
#pragma XLLEXPORT
	static XLOPERX x;

	/*
	long n = r*c;
	x = OPERX(n, new BYTE[n]);
	ensure (x.xltype == xltypeBigData);
	*/
	x.xltype = xltypeMulti;
	x.val.array.rows = r;
	x.val.array.columns = c;
	x.val.array.lparray = new XLOPERX[r*c];
	x.val.array.lparray[0].xltype = xltypeNum;
	x.val.array.lparray[0].val.num = r*c;

	// do stuff...

	x.xltype |= xlbitDLLFree;

	return &x;
}

static AddInX xai_xlfree_example(
	_T("?xll_xlfree_example"), XLL_LPXLOPERX XLL_WORDX XLL_WORDX,
	_T("XLFREE.EXAMPLE"), _T("rows, columns")
);
LPXLOPERX WINAPI
xll_xlfree_example(WORD r, WORD c)
{
#pragma XLLEXPORT
	static LOPERX o;

	o = ExcelX(xlUDF, ExcelX(xlfEvaluate, OPERX(_T("AUTO.EXAMPLE"))), OPERX(r), OPERX(c));

	// same as o.xltype |= xlbitXLFree; return &o;
	return o.XLFree();
}