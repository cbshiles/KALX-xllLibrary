// dropdown.cpp - macro for name-value pairs
#include "dropdown.h"

using namespace xll;

template<class X>
int xll_dropdown_macrox(void)
{
	// LISTBOX.PROPERTIES
	XOPER<X> sel = Excel<X>(xlfSelection);
	Excel<X>(xlcSelect, sel);
	XOPER<X> type = Excel<X>(xlfGetObject, XOPER<X>(50));
	XOPER<X> name = Excel<X>(xlfGetObject, XOPER<X>(10));
	XOPER<X> eval = XDropDown<X>::Evaluate(name);
	XOPER<X> ref = Excel<X>(xlfGetObject, XOPER<X>(4)); // upper left cell
	XOPER<X> item = Excel<X>(xlfGetObject, XOPER<X>(62)); // index of selected item
	XOPER<X> val = Excel<X>(xlfIndex, eval, item, XOPER<X>(eval.columns()));
	Excel<X>(xlfSetValue, ref, val);

	return 1;
}

static AddIn xai_dropdown_macro("_xll_dropdown_macro@0", "XLL.DROPDOWN.MACRO");
extern "C" int __declspec(dllexport) WINAPI
xll_dropdown_macro(void)
{
	return xll_dropdown_macrox<XLOPER>();
}
static AddIn12 xai_dropdown_macro12(L"_xll_dropdown_macro12@0", L"XLL.DROPDOWN.MACRO12");
extern "C" int __declspec(dllexport) WINAPI
xll_dropdown_macro12(void)
{
	return xll_dropdown_macrox<XLOPER12>();
}