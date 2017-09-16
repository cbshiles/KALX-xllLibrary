// on.h - xlcOnXXX functions
// Copyright (c) 2010 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include "xll/addin.h"

// used with On<Key>
#define ON_SHIFT   "+"
#define ON_CTRL    "^"
#define ON_ALT     "%"
#define ON_COMMAND "*"

namespace xll {

	struct Data {
		static const int On = xlcOnData;
	};
	struct Doubleclick {
		static const int On = xlcOnDoubleclick;
	};
	struct Entry {
		static const int On = xlcOnEntry;
	};
	struct Key {
		static const int On = xlcOnKey;
	};
	struct Recalc {
		static const int On = xlcOnRecalc;
	};
	struct Sheet {
		static const int On = xlcOnSheet;
	};
	struct Time {
		static const int On = xlcOnTime;
	};
	struct Window {
		static const int On = xlcOnWindow;
	};
/*
	// On<Key> ok("shortcut", "MACRO");
	template<class K>
	class On {
	public:
		typedef traits<XLOPERX>::xcstr xcstr;
		On(xcstr text, xcstr macro)
		{
			OPER x(3,1);

			x[0] = K::On;
			x[1] = text;
			x[2] = macro;

			Auto<Open> xao_on(x);
			Auto<Close> xac_on(x);
		}
		On(xcstr text, xcstr macro, bool activate)
		{
			OPER x(4,1);

			ensure (K::On == xlcOnSheet);
			x[0] = K::On;
			x[1] = text;
			x[2] = macro;
			x[3] = activate;

			Auto<Open> xao_on(x);
			Auto<Close> xac_on(x);
		}
		On(const OPER& time, xcstr macro, 
			const OPER& tolerance = MissingX(), bool insert = true)
		{
			OPER x(5,1);

			ensure (K::On == xlcOnTime);
			x[0] = K::On;
			x[1] = time;
			x[2] = macro;
			x[3] = tolerance;
			x[4] = insert;

			Auto<Open> xao_on(x);
			Auto<Close> xac_on(x);
		}
	};
*/
} // namespace xll
