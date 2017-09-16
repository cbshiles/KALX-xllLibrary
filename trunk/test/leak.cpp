// leak.cpp - try to leak
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"
#include "../xll/utility/srng.h"

using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xword xword;

utility::srng rng(false);

int
test_leak(void)
{
	try {
		OPERX o;

		o.resize((xword)rng.between(0, 200), (xword)rng.between(0, 20));
		for (xword i = 0; i < o.rows(); ++i)
			for (xword j = 0; j < o.columns(); ++j) {
				size_t len = (xchar)rng.between(0, 255);
				xchar* buf = new xchar[len + 1];
				buf[0] = (xchar)len;
				for (size_t k = 1; k <= len; ++k)
					buf[k] = (xchar)rng.between(0, 255);
				o(i, j) =  StrX(buf + 1, buf[0]);
				delete [] buf;
			}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfter> xao_leak(test_leak);