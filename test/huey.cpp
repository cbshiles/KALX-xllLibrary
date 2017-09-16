// huey.cpp - Baby Huey bangs on your spreadsheet
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"
#include "xll/xll.h"

using namespace xll;

static AddInX xai_baby_huey(_T("?xll_baby_huey"), _T("BABY.HUEY"));
int WINAPI
xll_baby_huey(void)
{
#pragma XLLEXPORT
	Excel<XLOPERX>(xlcAlert, StrX(_T("Bam on the Escape key to stop.")));

	while (!Excel<XLOPER>(xlAbort)) {
		Excel<XLOPER>(xlcCalculateNow);
	}

	return 1;
}