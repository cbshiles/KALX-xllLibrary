// auto.h - xlAutoXXX functions
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
//
// Can't Auto<Open> void WINAPI foo(void) functions.
// Can't Auto<Open> mem_fun's.
#pragma once
#include <list>

// Export well known functions.
#pragma comment(linker, "/include:_DllMain@12")
#pragma comment(linker, "/export:_XLCallVer@0")
#pragma comment(linker, "/export:xlAutoOpen@0=xlAutoOpen")
#pragma comment(linker, "/export:xlAutoClose@0=xlAutoClose")
#pragma comment(linker, "/export:xlAutoAdd@0=xlAutoAdd")
#pragma comment(linker, "/export:xlAutoRemove@0=xlAutoRemove")
//#pragma comment(linker, "/export:xlAutoFree@4=xlAutoFree")
//#pragma comment(linker, "/export:xlAutoFree12@4=xlAutoFree12")
//#pragma comment(linker, "/export:xlAutoRegister@4=xlAutoRegister")
//#pragma comment(linker, "/export:xlAutoRegister12@4=xlAutoRegister12")
//#pragma comment(linker, "/export:xlAddInManagerInfo@4=xlAddInManagerInfo")
//#pragma comment(linker, "/export:xlAddInManagerInfo12@4=xlAddInManagerInfo12")

#ifdef _DEBUG
#pragma comment(linker, "/export:?crtDbg@@3UCrtDbg@@A")
#endif

namespace xll {

	class Open {};
	class OpenAfter {};
	class Close {};
	class Add {};
	class RemoveBefore {};
	class Remove {};
	class Free {};
	class Free12 {};
	class Register {};
	class Register12 {};

	// Register macros to be called in xlAuto functions.
	template<class T>
	class Auto {
	public:
		Auto(int(*m)(void))
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
	private:
		typedef std::list<int(*)(void)> macro_list;
		typedef macro_list::iterator macro_iter;
		static macro_list& macros(void)
		{
			static std::list<int(*)(void)> macros_;

			return macros_;
		}
	};

} // namespace xll
