// on.cpp - OnXXX implementation
#include "xll/on.h"

using namespace xll;
/*
int 
xll::OnOpenX(const OPER& x)
{
	int f = static_cast<int>(x[0].val.num);

	if (x.size() == 3)
		Excel<XLOPER>(f, x[1], x[2]);
	else if (x.size() == 4)
		Excel<XLOPER>(f, x[1], x[2], x[3]);
	else if (x.size() == 5)
		Excel<XLOPER>(f, x[1], x[2], x[3], x[4]);

	ensure (x.size() < 6);

	return 1;
}
int
xll::OnCloseX(const OPER& x)
{
	if (x.size() > 1) {
		int f = static_cast<int>(x[0].val.num);
		Excel<XLOPER>(f, x[1]); // unregister
	}

	return 1;
}
*/