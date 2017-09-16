// auto.h - xlAutoXXX functions
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
//
// Can't Auto<Open> void WINAPI foo(void) functions.
// Can't Auto<Open> mem_fun's.
// Use std::bind???
#pragma once
#include <list>
#include <set>
#include "xll/oper.h"

// Export well known functions.
#pragma comment(linker, "/include:_DllMain@12")
#pragma comment(linker, "/export:_XLCallVer@0")
#pragma comment(linker, "/export:xlAutoOpen@0=xlAutoOpen")
#pragma comment(linker, "/export:xlAutoClose@0=xlAutoClose")
#pragma comment(linker, "/export:xlAutoAdd@0=xlAutoAdd")
#pragma comment(linker, "/export:xlAutoRemove@0=xlAutoRemove")
#pragma comment(linker, "/export:xlAutoFree@4=xlAutoFree")
#pragma comment(linker, "/export:xlAutoFree12@4=xlAutoFree12")
#pragma comment(linker, "/export:xlAutoRegister@4=xlAutoRegister")
#pragma comment(linker, "/export:xlAutoRegister12@4=xlAutoRegister12")
//#pragma comment(linker, "/export:xlAddInManagerInfo@4=xlAddInManagerInfo") // default is add-in name
//#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")

#ifdef _DEBUG
#pragma comment(linker, "/export:?crtDbg@@3UCrtDbg@@A")
#endif

namespace xll {

	class Open {};
	class Open12 {};
	class OpenAfter {};
	class OpenAfter12 {};
	class Close {};
	class Close12 {};
	class Add {};
	class Add12 {};
	class RemoveBefore {};
	class RemoveBefore12 {};
	class Remove {};
	class Remove12 {};
	class Free {};
	class Free12 {};

#ifndef EXCEL12
	typedef Open         OpenX;
	typedef OpenAfter    OpenAfterX;
	typedef Close        CloseX;
	typedef Add          AddX;
	typedef RemoveBefore RemoveBeforeX;
	typedef Remove       RemoveX;
	typedef Free         FreeX;
#else
	typedef Open12         OpenX;
	typedef OpenAfter12    OpenAfterX;
	typedef Close12        CloseX;
	typedef Add12          AddX;
	typedef RemoveBefore12 RemoveBeforeX;
	typedef Remove12       RemoveX;
	typedef Free12         FreeX;
#endif

	// Register macros to be called in xlAuto functions.
	template<class T>
	class Auto {
	public:
		typedef int (*macro)(void);
		Auto(macro m)
		{
			macros().push_back(m);
		}
		static int Call(void)
		{
			int result(1);

			macro_list& m(macros());
			for (macro_iter i = m.begin(); i != m.end(); ++i) {
				result = (*i)();
				if (!result)
					return result;
			}

			return result;
		}
/*
		Auto(const XOPER<XLOPER>& x)
		{
			callbacks().insert(x);
		}
		static void AutoOpen(void)
		{
			callback_list& m(callbacks());
			for (callback_iter i = m.begin(); i != m.end(); ++i) {
				const OPER& x = *i;

				int f = static_cast<int>(x[0].val.num);

				if (x.size() == 3)
					Excel<XLOPER>(f, x[1], x[2]);
				else if (x.size() == 4)
					Excel<XLOPER>(f, x[1], x[2], x[3]);
				else if (x.size() == 5)
					Excel<XLOPER>(f, x[1], x[2], x[3], x[4]);

				ensure (x.size() < 6);
			}
		}
		static void AutoClose(void)
		{
			callback_list& m(callbacks());
			for (callback_iter i = m.begin(); i != m.end(); ++i) {
				OPER& x = *i;

				ensure (x.size() > 1);
				int f = static_cast<int>(x[0].val.num);

				Excel<XLOPER>(f, x[1]);
			}
		}
*/	private:
		typedef std::list<macro> macro_list;
		typedef macro_list::iterator macro_iter;
		static macro_list& macros(void)
		{
			static macro_list macros_;

			return macros_;
		}
/*
		typedef std::set<XOPER<XLOPER>> callback_list;
		typedef typename callback_list::iterator callback_iter;
		static callback_list& callbacks(void)
		{
			static callback_list callbacks_;

			return callbacks_;
		}
*/	};

	template<>
	class Auto<Free> {
	public:
		typedef void (*macro)(LPXLOPER);
		Auto(macro m)
		{
			macros().push_back(m);
		}
		static int Call(LPXLOPER px)
		{
			int result(1);

			macro_list& m(macros());
			for (macro_iter i = m.begin(); i != m.end(); ++i) {
				(*i)(px);
				if (!result)
					return result;
			}

			return result;
		}

	private:
		typedef std::list<macro> macro_list;
		typedef macro_list::iterator macro_iter;
		static macro_list& macros(void)
		{
			static macro_list macros_;

			return macros_;
		}
	}; 
	
	template<>
	class Auto<Free12> {
	public:
		typedef void (*macro)(LPXLOPER12);
		Auto(macro m)
		{
			macros().push_back(m);
		}
		static int Call(LPXLOPER12 px)
		{
			int result(1);

			macro_list& m(macros());
			for (macro_iter i = m.begin(); i != m.end(); ++i) {
				(*i)(px);
				if (!result)
					return result;
			}

			return result;
		}

	private:
		typedef std::list<macro> macro_list;
		typedef macro_list::iterator macro_iter;
		static macro_list& macros(void)
		{
			static macro_list macros_;

			return macros_;
		}
	}; 
} // namespace xll
