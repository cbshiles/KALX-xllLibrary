// test.cpp - Exercise the xll add-in library.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"

using namespace xll;

static AddInX xai_locus(
	_T("?xll_locus"), XLL_LPOPERX XLL_UNCALCEDX XLL_BOOLX,
	_T("TEST.LOCUS"), _T("Reset?"),
	_T("Test the locus class.")
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
