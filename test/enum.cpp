// enum.cpp - enumerated values
// Convert string to enumeration value.
// e.g., FREQ("SemiAn") = 2
// FREQ.SemiAnnually() = 2
#include <algorithm>
#include "xll/xll.h"

using namespace xll;

// deprecate???
static Enum freq(Name("FREQ")
	.Item("Annually", 1, "once per year")
	.Item("SemiAnnually", 2, "twice per year")
	.Item("Quarterly", 4, "four times per year")
	.Item("Monthly", 12, "twelve times per year")
	.Category("Enumerations")
);

int xll_test_enum(void)
{
	try {
		OPERX o;
		o = Enum::Value(_T("FREQ"), _T("Ann"));
		ensure (o == 1);

		o = Enum::Value(_T("FREQ"), OPERX(123));
		ensure (o == 123); // numbers are passed through

		ensure (Enum::Value(_T("FREQ"), _T("m")) == 12);
		ensure (Enum::value(_T("FREQ"), _T("q")) == 4);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_enum(xll_test_enum);