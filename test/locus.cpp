// test.cpp - Exercise the xll add-in library.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"

using namespace xll;

static AddInX xai_locus(
	_T("?xll_locus"), XLL_LPOPERX XLL_UNCALCEDX XLL_BOOLX,
	_T("TEST.LOCUS"), _T("Reset?"),
	_T("LOCUS"), _T("Test the locus class.")
);
LPOPERX WINAPI
xll_locus(BOOL reset)
{
#pragma XLLEXPORT
	static OPERX xResult;
	locus cell;

	if (reset) {
		cell.set(NumX(0));
		xResult = 0;
	}
	else {
		xResult = cell.get();
		xResult = xResult + 1;
		cell.set(xResult);
	}

	return &xResult;
}

static AddInX xai_test_name(_T("?xll_test_name"), _T("TEST.NAME"));
int WINAPI
xll_test_name(void)
{
#pragma XLLEXPORT
	OPERX o;

	o = ExcelX(xlfSetName, OPERX(_T("myname")), OPERX(123));
	ensure (o == true);
	o = ExcelX(xlfGetName, OPERX(_T("myname")));
	ensure (o.xltype == xltypeStr);
	ensure (o == _T("=123"));

	return 1;
}

static AddInX xai_all_names(_T("?xll_all_names"), _T("ALL.NAMES"));
int WINAPI
xll_all_names(void)
{
#pragma XLLEXPORT

	ExcelX(xlcAlert, locus::Names()); // no names for xlcSetName!!!

	return 1;
}