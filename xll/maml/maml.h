// maml.h - Microsoft Awful Markup Language
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "xml.h"
#define WINDOWS_LEAN_AND_MEAN
#include <windows.h>
#include <tchar.h>

// create Sandcastle documentation
#ifndef _LIB
#ifdef _DEBUG
#pragma comment(linker, "/include:_xll_make_doc@0")
#endif // _DEBUG
#endif // _LIB

namespace maml {

	// phony guid using topic id
	inline LPCTSTR 
	TUID(unsigned int tid)
	{
		static TCHAR guid[64];

		_stprintf_s(guid, sizeof(guid)/sizeof(TCHAR), _T("%04x%04x-0000-0000-0000-000000000000"), tid>>16, tid&0xFFFF);
		
		return guid;
	}


} // namespace maml