// function.cpp - Rename this file and replace this description.
#include "header.h"

using namespace xll;

static AddInX xai_function(
	// Return type, C name, Excel name.
	FunctionX(XLL_LPOPERX, _T("?xll_function"), _T("XLL.FUNCTION"))
	// Type, name, description
	.Arg(XLL_DOUBLEX, _T("Number"), _T("is a number. "))
//	.Arg(...) // add more args here
	.Category(CATEGORY)
	.FunctionHelp(_T("Description of what function does."))
	.Documentation(_T("Optional documentation for function."))
);
LPOPERX WINAPI // <- Must declare all Excel C functions as WINAPI
xll_function(double x)
{
#pragma XLLEXPORT // <- Don't forget this or your function won't get exported.
	static OPERX xResult;

	// function body goes here
	xResult = x; // Wait, what? Yes, an OPER can act like a double!
	
	return &xResult;
}

static AddInX xai_sampler(
	FunctionX(XLL_LPOPERX, _T("?xll_sampler"), _T("XLL.SAMPLER"))
	// See xll/defines.h for a complete list.
	.Arg(XLL_BOOLX, 
		_T("Bool"), _T("is a short int used as a logical."))
	.Arg(XLL_DOUBLEX, 
		_T("Num"), _T("is a 64-bit IEEE double."))
	.Arg(XLL_CSTRINGX, 
		_T("Str"), _T("is a char* NULL terminated C style string."))
	.Arg(XLL_PSTRINGX, 
		_T("Str"), _T("is an unsigned char* Pascal style string with the first char being the count."))
	.Arg(XLL_WORDX, 
		_T("Int"), _T("is an unsigned 2 byte int."))
	.Arg(XLL_SHORTX, 
		_T("Int"), _T("is a signed 2 byte int."))
	.Arg(XLL_LONGX, 
		_T("Int"), _T("is a signed 4 byte int."))
	.Arg(XLL_FPX, 
		_T("Array"), _T("is a pointer to a two dimensional array of floating point doubles."))
	.Arg(XLL_LPOPERX, 
		_T("Range"), _T("is a pointer to an OPER struct (never a reference type)."))
	.Arg(XLL_LPXLOPERX, 
		_T("Range"), _T("is a pointer to an XLOPER struct. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Sampler of the the most common argument types."))
	.Documentation(_T("Returns the first argument that is not zero/null/missing."))
);
LPOPERX WINAPI
xll_sampler(BOOL b, double x, const char* s, xcstr p, WORD w, SHORT i, LONG l, _FP* pa, LPOPERX po, LPXLOPERX px)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		if (b) {
			o = b;
			ensure (o.xltype == xltypeNum); // Excel converts ints to doubles.
			o = b ? true : false;
			ensure (o.xltype == xltypeBool);
			ensure (o.val.xbool == b);
		}
		else if (x) {
			o = x;
			ensure (o.xltype == xltypeNum);
			ensure (o.val.num == x);
		}
		else if (s && *s) {
			o = s;
			ensure (o == s);
		}
		else if (p && *p) {
			o = OPERX(p + 1, p[0]); // counted string constructor
		}
		else if (w) {
			o = w;
			ensure (o.xltype == xltypeNum);
			o = IntX(w); // if you really need an int
			ensure (o.xltype == xltypeInt);
			ensure (o.val.w == w);
		}
		else if (i) {
			o = i;
			ensure (o.xltype == xltypeNum);
			ensure (o.val.num == i);
		}
		else if (l) {
			o = l;
			ensure (o.xltype == xltypeNum);
			ensure (o.val.num == l);
		}
		else if (pa->array[0]) {
			// OPER's cannot be FP types!
			FPX a = *pa;
			o.resize(pa->rows, pa->columns);
			for (xword i = 0; i < pa->rows; ++i) {
				for (xword j = 0; j < pa->columns; ++j) {
					o(i, j) = a(i, j);
					ensure (o(i, j).xltype == xltypeNum);
				}
			}
		}
		else if (*po) {
			o = *po;
			ensure (o.xltype == po->xltype);
			ensure (o.val.array.rows == po->val.array.rows);
			ensure (o.val.array.columns == po->val.array.columns);
		}
		else if (px->xltype != xltypeMissing) { // raw struct
			o = *px;
			ensure (o.xltype == px->xltype);
			ensure (o.val.array.rows == px->val.array.rows);
			ensure (o.val.array.columns == px->val.array.columns);
		}
		else {
			XLL_INFO("XLL.SAMPLER: no nontrivial arguments.", true); // force info popup
		}
	}
	catch (const std::exception& ex) {
		XLL_WARNING(ex.what());

		o = ErrX(xlerrValue);
	}

	return &o;
}