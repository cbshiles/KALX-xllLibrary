// traits.h - Parameterize by XLOPER and XLOPER12.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "xll/limits.h"

template<class X> class XOPER;

namespace xll {

	template<class X>
	struct traits { };

	template<>
	struct traits<XLOPER> {
		typedef XLOPER type;
		typedef XLOPER* pointer_type;
		typedef unsigned short xbool;
		typedef short int xint;
		typedef unsigned short xword;
		typedef char xchar;
		typedef const char* xcstr;
		typedef _FP xfp;
		typedef WORD xrw;
		typedef BYTE xcol;
		typedef std::basic_string<xchar> xstring;

		static const int strmax = 255; // maximum number of characters in a str

		static int Excel(int xlfn, LPXLOPER operRes, int count, ...)
		{
			LPXLOPER arg[limits<XLOPER>::maxargs];

			ensure (count <= limits<XLOPER>::maxargs);

			va_list ap;
			va_start(ap, count);
			for (int i = 0; i < count; ++i)
				arg[i] = va_arg(ap, LPXLOPER);
			va_end(ap);

			return Excel4v(xlfn, operRes, count, &arg[0]);
		}
		static int Excelv(int xlfn, LPXLOPER operRes, int count, const LPXLOPER opers[])
		{
			return Excel4v(xlfn, operRes, count, const_cast<LPXLOPER*>(opers));
		}

		static xchar* strncpy(xchar* dest, xcstr src, xchar n)
		{
			return ::strncpy(dest, src, n);
		}
		static int strnicmp(xcstr s, xcstr t, xchar n)
		{
			return ::_strnicmp(s, t, n);
		}
		static xchar strlen(xcstr s)
		{
			size_t n = ::strlen(s);
			ensure (n <= limits<XLOPER>::maxchars);

			return static_cast<xchar>(n);
		}
		static xcstr null(void)
		{
			return "";
		}
		static xcstr chm(void)
		{
			return "chm";
		}
		static const XLOPER& format(void)
		{
			static XLOPER fmt;

			fmt.xltype = xltypeStr;
			fmt.val.str = "1@";

			return fmt;
		}
	};

	template<>
	struct traits<XLOPER12> {
		typedef XLOPER12 type;
		typedef XLOPER12* pointer_type;
		typedef int xbool;
		typedef int xint;
		typedef unsigned long xword;
		typedef wchar_t xchar;
		typedef const wchar_t* xcstr;
		typedef _FP12 xfp;
		typedef RW xrw;
		typedef COL xcol;
		typedef std::basic_string<xchar> xstring;

		static const int strmax = 32767;

		static int Excel(int xlfn, LPXLOPER12 operRes, int count, ...)
		{
			LPXLOPER12 arg[limits<XLOPER12>::maxargs];

			ensure (count <= limits<XLOPER12>::maxargs);

			va_list ap;
			va_start(ap, count);
			for (int i = 0; i < count; ++i)
				arg[i] = va_arg(ap, LPXLOPER12);
			va_end(ap);

			return Excel12v(xlfn, operRes, count, &arg[0]);
		}
		static int Excelv(int xlfn, LPXLOPER12 operRes, int count, const LPXLOPER12 opers[])
		{
			return Excel12v(xlfn, operRes, count, const_cast<LPXLOPER12*>(opers));
		}

		static xchar* strncpy(xchar* dest, xcstr src, xchar n)
		{
			return ::wcsncpy(dest, src, n);
		}
		static int strnicmp(xcstr s, xcstr t, xchar n)
		{
			return ::_wcsnicmp(s, t, n);
		}
		static xchar strlen(xcstr s)
		{
			size_t n = ::wcslen(s);
			ensure (n <= limits<XLOPER12>::maxchars);

			return static_cast<xchar>(n);
		}
		static xcstr null(void)
		{
			return L"";
		}
		static xcstr chm(void)
		{
			return L"chm";
		}
		static const XLOPER12& format(void)
		{
			static XLOPER12 fmt;

			fmt.xltype = xltypeStr;
			fmt.val.str = L"1@";

			return fmt;
		}
	};

} // namespance xll