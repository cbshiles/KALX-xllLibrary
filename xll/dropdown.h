// dropdown.h - linked dropdown boxes
// Copyright (c) 2012 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <list>
#include "addin.h"

#ifdef _M_X64 
#pragma comment(linker, "/include:xll_dropdown_macro")
#pragma comment(linker, "/include:xll_dropdown_macro12")
#else
#pragma comment(linker, "/include:_xll_dropdown_macro@0")
#pragma comment(linker, "/include:_xll_dropdown_macro12@0")
#endif

namespace xll {
	template<class X>
	struct dropdown_traits {
		static XOPER<X> macro(void);
	};
	template<>
	struct dropdown_traits<XLOPER> {
		static OPER macro(void) { return OPER("XLL.DROPDOWN.MACRO"); }
	};
	template<>
	struct dropdown_traits<XLOPER12> {
		static OPER12 macro(void) { return OPER12("XLL.DROPDOWN.MACRO12"); }
	};

	template<class X>
	class XDropDown {
		XOPER<X> name_;
		XOPER<X> o_; // two column array of name-value pairs
	public:
		XDropDown(XOPER<X> o, XOPER<X> ref1 = XMissing<X>(), XOPER<X> ref2 = XMissing<X>())
			: o_(o)
		{
			if (ref1.xltype == xltypeMissing) {
				ref1 = Excel<X>(xlfSelection);
			}
			if (ref2.xltype == xltypeMissing) {
				ref2 = Excel<X>(xlfOffset, ref1, XOPER<X>(1), XOPER<X>(1));
			}

			name_ = Excel<X>(xlfCreateObject, XOPER<X>(20), ref1, XOPER<X>(0), XOPER<X>(0), ref2);
			o = Evaluate(o);
			for (xword i = 0; i < o.rows(); ++i) {
				ExcelX(xlcAddListItem, o(i, 0));
			}
			XOPER<X> xRef = Excel<X>(xlfReftext, ref1, XOPER<X>(true));
			Excel<X>(xlcListboxProperties, XMissing<X>(), xRef);
		}
		~XDropDown()
		{
//			Select();
//			Excel<X>(xlcCut);
		}
		void Assign(const XOPER<X>& macro)
		{
			Select();
			Excel<X>(xlcAssignToObject, macro);
		}
		void Hide(bool b)
		{
			Excel<X>(xlcHideObject, XOPER<X>(b));
		}
		const XOPER<X>& Name(void) const
		{
			return name_;
		}
		void Rename(const XOPER<X>& name)
		{
			Select();
			ensure (Excel<X>(xlcRenameObject, name));
			name_ = name;
		}
		void Select() const 
		{
			Excel<X>(xlcSelect, name_);
		}
		XOPER<X> Enum(void) const
		{
			return evaluate(name_);
		}
		static XOPER<X> Evaluate(const XOPER<X>& o)
		{
			static traits<X>::xchar eq('='), lp('('), rp(')');

			return Excel<X>(xlfEvaluate, Excel<X>(xlfConcatenate, XOPER<X>(&eq,1), o, XOPER<X>(&lp,1), XOPER<X>(&rp, 1)));
		}
	};

	typedef XDropDown<XLOPER>	DropDown;
	typedef XDropDown<XLOPER12> DropDown12;
	typedef XDropDown<XLOPERX>	DropDownX;

} // namespace xll