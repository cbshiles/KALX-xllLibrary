// error.h - Error reporting functions
// Copyright (c) 2005-2011 KALX, LLC. All rights reserved.
#pragma once

// error reporting macros
#ifndef _LIB
#pragma comment(linker, "/include:_xll_alert_filter@0")
#endif // _LIB

enum { 
	XLL_ALERT_ERROR   = 1,
	XLL_ALERT_WARNING = 2, 
	XLL_ALERT_INFO    = 4,
	XLL_ALERT_LOG     = 8	// turn on logging (not yet implemented)
};

int XLL_ERROR(const char* e, bool force = false);
int XLL_WARNING(const char* e, bool force = false);
int XLL_INFORMATION(const char* e, bool force = false);
