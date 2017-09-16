// auto.cpp - Implement well-known xlAutoXXX entry points required for xll's.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "xll.h"

using namespace xll;

extern "C" 
int WINAPI
xlAutoOpen(void)
{
	try {
//		Auto<Open>::AutoOpen();
		ensure (Auto<Open>::Call());
		ensure (AddIn::RegisterAll());
		ensure (Auto<OpenAfter>::Call());

		if (XLCallVer() >= 0x0C00) {
			ensure (Auto<Open12>::Call());
			ensure (AddIn12::RegisterAll());
			ensure (Auto<OpenAfter12>::Call());
		}
		
	}
	catch (const std::exception& ex) {
		Excel<XLOPER>(xlcAlert, Str(ex.what()));

		return 0;
	}

	return 1;
}

extern "C"
int WINAPI
xlAutoClose(void)
{
	try {
		ensure (Auto<Close>::Call());
//		Auto<Close>::AutoClose();
		if (XLCallVer() >= 0x0C00) {
			ensure (Auto<Close12>::Call());
		}
	}
	catch (const std::exception& ex) {
		Excel<XLOPER>(xlcAlert, Str(ex.what()));

		return 0;
	}

	return 1;
}

extern "C"
int WINAPI
xlAutoAdd(void)
{
	try {
		ensure (Auto<Add>::Call());
		if (XLCallVer() >= 0x0C00) {
			ensure (Auto<Add12>::Call());
		}
	}
	catch (const std::exception& ex) {
		Excel<XLOPER>(xlcAlert, Str(ex.what()));

		return 0;
	}

	return 1;
}

extern "C"
int WINAPI
xlAutoRemove(void)
{
	try {
		ensure (Auto<RemoveBefore>::Call());
		ensure (AddIn::UnregisterAll());
		ensure (Auto<Remove>::Call());

		if (XLCallVer() >= 0x0C00) {
			ensure (Auto<RemoveBefore12>::Call());
			ensure (AddIn12::UnregisterAll());
			ensure (Auto<Remove12>::Call());
		}

	}
	catch (const std::exception& ex) {
		Excel<XLOPER>(xlcAlert, Str(ex.what()));

		return 0;
	}

	return 1;
}

extern "C"
void WINAPI
xlAutoFree(LPXLOPER px)
{
	Auto<Free>::Call(px);
}

extern "C"
void WINAPI
xlAutoFree12(LPXLOPER12 px)
{
	Auto<Free12>::Call(px);
}

extern "C"
LPXLOPER WINAPI 
xlAutoRegister(LPXLOPER pxName)
{
	static OPER xResult;

	try {
		const Args* parg = ArgsMap::Find(*pxName);

		if (parg) {
			xResult = xll::Register(*parg);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		xResult = Err(xlerrValue);
	}

	return &xResult;
}

extern "C"
LPXLOPER12 WINAPI 
xlAutoRegister12(LPXLOPER12 pxName)
{
	static OPER12 xResult;

	try {
		const Args12* parg = ArgsMap12::Find(*pxName);

		if (parg) {
			xResult = xll::Register(*parg);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		xResult = Err12(xlerrValue);
	}

	return &xResult;
}