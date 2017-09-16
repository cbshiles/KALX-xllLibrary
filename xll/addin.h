// addin.h - Excel add-in data and registration.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "register.h"

namespace xll {

	template<class X>
	class XAddIn {
		XArgs<X> arg_; // arguments required to register the add-in
		double regid_; // really an int
		FARPROC proc_; // underlying C function		
	public:
		typedef typename xll::traits<X>::xcstr xcstr;
		typedef typename xll::traits<X>::xword xword;
		// macro
		XAddIn(xcstr procedure, xcstr function_text)
			: arg_(procedure, function_text)
		{
			List().push_back(this);
		}
		// function
		XAddIn(xcstr procedure, xcstr type_text, xcstr function_text, xcstr argument_text,
			xcstr category = 0, xcstr function_help = 0,int count = 0, xcstr* argument_help = 0,
			xcstr documentation = 0)
			: arg_(procedure, type_text, function_text, argument_text)
		{
			if (category)
				arg_.Category(category);
			if (function_help)
				arg_.FunctionHelp(function_help);
			for (int i = 0; i < count; ++i)
				arg_.ArgumentHelp(argument_help[i]);
			if (documentation)
				arg_.Documentation(documentation);

			List().push_back(this);
		}
		// general case
		XAddIn(const XArgs<X>& arg)
			: arg_(arg)
		{
			List().push_back(this);
		}

		const XArgs<X>& Arg(void) const
		{
			return arg_; // arguments required to register the add-in
		}

		int RegisterId(void) const
		{
			return static_cast<int>(regid_);
		}
		FARPROC GetProc(void) const
		{
			reutrn proc_;
		}

		XOPER<X> Register()
		{
			XOPER<X> oReg = xll::Register(arg_);
			proc_ = (arg_.isFunction() || arg_.isMacro()) ? GetProcAddress(GetModule(), GetProc()) : 0;
			if (oReg.xltype == xltypeNum)
				regid_ = oReg.val.num;

			return oReg;
		}
		XOPER<X> Unregister()
		{
			regid_ = 0;
			
			return xll::Unregister(arg_);
		}

		static int RegisterAll()
		{
			addin_list l(List());

			for (addin_citer i = l.begin(); i != l.end(); ++i) {
				if (!((*i)->Register()))
					return 0;
			}

			return 1;
		}
		static int UnregisterAll()
		{
			const addin_list& l(List());

			for (addin_citer i = l.begin(); i != l.end(); ++i) {
				if (!((*i)->Arg().isDocument()))
					if (!((*i)->Unregister()))
						return 0;
			}

			return 1;
		}

		static LPCSTR GetName(void)
		{
			static OPER oGetName(Excel<XLOPER>(xlGetName));
			static std::string name_(oGetName.val.str + 1, oGetName.val.str[0]);

			return name_.c_str();
		}
		static HMODULE GetModule(void)
		{
			static HMODULE module_(GetModuleHandle(GetName()));

			return module_;
		}
		LPCSTR GetProc(void)
		{
			static std::string proc;
			XOPER<X> oProc = arg_.Procedure();

			proc.erase();
			for (xword i = 1; i <= oProc.val.str[0]; ++i)
				proc.push_back(static_cast<CHAR>(oProc.val.str[i]));
			proc.push_back(0);

			return proc.c_str();
		}

		typedef std::list<XAddIn<X>*> addin_list;
		typedef typename addin_list::iterator addin_iter;
		typedef typename addin_list::const_iterator addin_citer;
		static addin_list& List() 
		{
			static addin_list l_;

			return l_;
		}
	};

	typedef XAddIn<XLOPER>   AddIn;
	typedef XAddIn<XLOPER12> AddIn12;
	typedef X_(AddIn)        AddInX;

} // namespace xll