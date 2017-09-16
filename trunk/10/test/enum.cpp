// enum.cpp - enumerated values
#include "xll/xll.h"

using namespace xll;

static Enum e(EnumName("Foo_Enum")
	.Item("Foo_One", 1, "is the number one")
	.Item("Foo_Two", 2, "is the number two")
	.Category("EnumCat")
);

static AddIn xai_enum_test(
	Function(XLL_LONG, "?xll_enum_test", "ENUM.TEST")
	.Arg(XLL_LPOPER, "Enum", "is an enum")
);
LONG WINAPI
xll_enum_test(LPOPER pe)
{
#pragma XLLEXPORT
	int i(-1);

	try {
		i = Enum::Value("Foo_Enum", *pe);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return i;
}

static AddInX xai_enum(
	_T("?xll_enum"), XLL_LPOPERX XLL_LPOPERX XLL_UNCALCEDX,
	_T("ENUM2"), _T("")
);
LPOPERX WINAPI
xll_enum(LPOPER oe)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o = ExcelX(xlfEvaluate, *oe);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &o;
}
