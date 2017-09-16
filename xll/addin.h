// addin.h - Excel add-in data and registration.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "xll/register.h"

namespace xll {

	template<class X>
	class XAddIn {
		typedef typename xll::traits<X>::xcstr xcstr;

		XArgs<X> arg_; // arguments required to register the add-in
	public:
		// macro
		XAddIn(xcstr procedure, xcstr function_text)
			: arg_(procedure, function_text)
		{
			List().push_back(&arg_);
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

			List().push_back(&arg_);
		}
		// general case
		XAddIn(const ArgsX& arg)
			: arg_(arg)
		{
			List().push_back(&arg_);
		}

		const ArgsX& Arg(void) const
		{
			return arg_; // arguments required to register the add-in
		}

		static int Register()
		{
			addin_list l(List());

			for (addin_citer i = l.begin(); i != l.end(); ++i) {
				if (!AddInRegister<X>(*(*i)))
					return 0;
			}

			return 1;
		}
		static int Unregister()
		{
			const addin_list& l(List());

			for (addin_citer i = l.begin(); i != l.end(); ++i) {
				if (!AddInUnregister<X>(*(*i)))
					return 0;
			}

			return 1;
		}

		typedef std::list<XArgs<X>*> addin_list;
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