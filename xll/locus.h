// locus.h - location in a spreadsheet
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <cctype>
#include "oper.h"

namespace xll {

	inline void
	text2name(traits<XLOPERX>::xchar* t)
	{
		for (int i = 1; i <= t[0]; ++i)
			if (!isalnum(t[i]))
				t[i] = _T('_');
	}
	inline void
	interchange(traits<XLOPERX>::xchar* t, traits<XLOPERX>::xchar r, traits<XLOPERX>::xchar c)
	{
		for (int i = 1; i < t[0]; ++i) {
			if (t[i] == r)
				t[i] = c;
			else if (t[i] == c)
				t[i] = r;
		}
	}

	// associate OPER with a spreadsheet location
	class locus {
		OPERX x_, y_;
	public:
		// locus l; // in uncalced function
		locus()
		{
			OPERX  oCaller; // possibly the old locus
			LOPERX xCaller = ExcelX(xlfCaller);
			LOPERX xReftext = ExcelX(xlfReftext, xCaller);

			x_ = ExcelX(xlfGetDef, xReftext);

			// not defined or caller is the active cell
			if (x_.xltype != xltypeStr || (xCaller == ExcelX(xlfActiveCell) && !ExcelX(xlCoerce, xCaller))) {
				x_ = xReftext;
				text2name(x_.val.str);
				// check for existing locus
				oCaller = ExcelX(xlfGetName, x_);
				// Excel tracks caller.
				ExcelX(xlfSetName, x_, xCaller);
			}

			ensure (x_.xltype == xltypeStr);

			// Name of actual value.
			y_ = x_;
			interchange(y_.val.str, _T('R'), _T('C'));

			// if old locus exists hook it up
			if (oCaller) {
				// strip off leading '='
				oCaller = ExcelX(xlfRight, oCaller, OPERX(oCaller.val.str[0] - 1));
				OPERX o_ = oCaller;
				text2name(o_.val.str);
				ExcelX(xlfSetName, o_, ExcelX(xlfTextref, oCaller));
				interchange(o_.val.str, _T('R'), _T('C'));
				ExcelX(xlfSetName, o_, ExcelX(xlfEvaluate, ExcelX(xlfGetName, y_)));
				ExcelX(xlfSetName, y_); // undefine existing name
			}
		}
		~locus()
		{ }
		OPERX get()
		{
			return ExcelX(xlfEvaluate, ExcelX(xlfGetName, y_));
		}
		void set(const XLOPERX& v)
		{
			ExcelX(xlfSetName, y_, v);
		}
	};

} // namespace xll