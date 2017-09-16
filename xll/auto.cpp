// auto.cpp - Implement well-known xlAutoXXX entry points required for xll's.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "xll/xll.h"
#include "xll/auto.h"

using namespace xll;

extern "C" 
int WINAPI
xlAutoOpen(void)
{
	try {
		ensure (Auto<Open>::Call());

		ensure (AddIn::Register());
		if (XLCallVer() >= 0x0C00)
			ensure (AddIn12::Register());
		
		ensure (Auto<OpenAfter>::Call());
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

		if (XLCallVer() >= 0x0C00)
			ensure (AddIn12::Unregister());
		ensure (AddIn::Unregister());

		ensure (Auto<Remove>::Call());
	}
	catch (const std::exception& ex) {
		Excel<XLOPER>(xlcAlert, Str(ex.what()));

		return 0;
	}

	return 1;
}