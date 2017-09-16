// async.cpp - Asynchronous function example.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
// http://msdn.microsoft.com/en-us/library/ff796219.aspx
// Put =XLL.ECHO(RAND()) in one cell and =RAND() in another.
// Hit F9 and note both cells update after 1 second.
// Put =XLL.ECHOA(RAND()) in one cell and =RAND() in another.
// Hit F9 and note =RAND() updates immediately.
#include "test.h"

using namespace xll;

// slow synchronous function
static AddIn12 xai_echo(
	L"?xll_echo", XLL_LPOPER12 XLL_LPOPER12,
	L"XLL.ECHO", L"OPER"
);
LPOPER12 WINAPI
xll_echo(LPOPER12 po)
{
#pragma XLLEXPORT
	Sleep(1000);

	if (po->xltype & xltypeNum)
		po->val.num *= 2;

	return po;
}

// using asynchronous facilities

#include <process.h>

void __cdecl xll_echow(void*); // forward declaration

// start a thread from Excel
static AddIn12 xai_echoa(
	L"?xll_echoa", XLL_VOID12 XLL_LPOPER12 XLL_ASYNCHRONOUS12 XLL_THREAD_SAFE12,
	L"XLL.ECHOA", L"OPER"
);
void WINAPI
xll_echoa(LPOPER12 pd, LPXLOPER12 ph)
{
#pragma XLLEXPORT
	OPER12& dh = *new OPER12(2, 1);

	dh[0] = *pd;
	// ensure (ph->xltype == xltypeBigData);
	dh[1] = *ph;

//	if (!CreateThread(0, 0, xll_echow, &dh, 0, 0)) { // use _begin/_endthread???
	if (-1L == _beginthread(xll_echow, 0, &dh)) { 
		// xll_echow not called
		delete &dh;
	}
}

// do the work
void __cdecl 
xll_echow(void* arg)
{
	OPER12& dh = *(LPOPER12)arg;

	Sleep(1000);

	if (dh[0].xltype & xltypeNum)
		dh[0].val.num *= 2;

	// return result to xll_echoa
	Excel<XLOPER12>(xlAsyncReturn, dh[1], dh[0]); // note handle, then data

	delete &dh;
	_endthread();
}
