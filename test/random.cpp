// random.cpp - generate random OPER's.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"
#include <ctime>
#include "xll/registry.h"
#include "xll/xll.h"

#ifndef CATEGORY
#define CATEGORY _T("Test")
#endif

using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xword xword;
typedef traits<XLOPERX>::xstring xstring;

#define SUBKEY _T("Software\\KALX\\xll")

// save seed to registry
struct seed {
	// use value in registry if true
	seed(bool read = false)
	{
		if (read) {
			Reg::Object<TCHAR, DWORD> s(HKEY_CURRENT_USER, SUBKEY, _T("seed"), 12345679);

			::srand(static_cast<unsigned int>(s));
		}
		else {
			DWORD s = static_cast<DWORD>(::time(0));

			::srand(static_cast<unsigned int>(s));

			Reg::CreateKey<TCHAR>(HKEY_CURRENT_USER, SUBKEY).SetValue<DWORD>(_T("seed"), s);
		}
	}
};

static seed s(false); // change to true to read from registry

inline int 
rand_between(int a, int b)
{
	return a + ((b - a)*rand())/(RAND_MAX + 1);
}

inline double
rand_uniform(double a = 0, double b = 1)
{
	return a + (b - a)*rand()/(RAND_MAX + 1);
}

XLL_ENUM(xltypeNum, TYPE_NUM, CATEGORY, _T("A 64-bit IEEE floating point number. "));
XLL_ENUM(xltypeStr, TYPE_STR, CATEGORY, _T("A character string. "));
XLL_ENUM(xltypeBool, TYPE_BOOL, CATEGORY, _T("A Boolean value. "));
XLL_ENUM(xltypeRef, TYPE_REF, CATEGORY, _T(" "));
XLL_ENUM(xltypeErr, TYPE_ERR, CATEGORY, _T("An error type. "));
XLL_ENUM(xltypeMulti, TYPE_MULTI, CATEGORY, _T("A two dimensional Range of cells. "));
XLL_ENUM(xltypeMissing, TYPE_MISSING, CATEGORY, _T("A missing type. "));
XLL_ENUM(xltypeNil, TYPE_NIL, CATEGORY, _T("A nil type. "));
XLL_ENUM(xltypeSRef, TYPE_SREF, CATEGORY, _T("A reference to a single range. "));
XLL_ENUM(xltypeInt, TYPE_INT, CATEGORY, _T("A 16-bit signed integer. "));

static AddInX xai_rand_num(
	_T("?xll_rand_num"), XLL_DOUBLEX XLL_VOLATILEX,
	_T("RAND.NUM"), _T(""),
	_T("Test"), _T("Return a random double. ")
);
double WINAPI
xll_rand_num(void)
{
#pragma XLLEXPORT

	return (rand_uniform() < 0.5 ? 1 : -1 )/rand_uniform();
}

static AddInX xai_rand_str(
	_T("?xll_rand_str"), XLL_PSTRINGX XLL_VOLATILEX,
	_T("RAND.STR"), _T(""),
	_T("Test"), _T("Return a random string. ")
);
const xchar* WINAPI
xll_rand_str(void)
{
#pragma XLLEXPORT
	static xstring s;

	xchar len = static_cast<xchar>(rand_between(0, limits<XLOPERX>::maxchars));
	s.resize(len + 1);
	
	s[0] = len;
	for (xchar i = 1; i <= len; ++i) {
		s[i] = static_cast<xchar>(rand_between(0, limits<XLOPERX>::maxchars));
	}

	return &s[0];
}

static AddInX xai_rand_bool(
	_T("?xll_rand_bool"), XLL_BOOLX XLL_VOLATILEX,
	_T("RAND.BOOL"), _T(""),
	_T("Test"), _T("Return a random boolean. ")
);
BOOL WINAPI
xll_rand_bool(void)
{
#pragma XLLEXPORT

	return rand_uniform() < 0.5 ? FALSE : TRUE;
}

static int err[] = {
    xlerrNull,
    xlerrDiv0,
    xlerrValue,
    xlerrRef,
    xlerrName,
    xlerrNum,
    xlerrNA,
    xlerrGettingData
};

static AddInX xai_rand_err(
	_T("?xll_rand_err"), XLL_LPOPERX XLL_VOLATILEX,
	_T("RAND.ERR"), _T(""),
	_T("Test"), _T("Return a random error. ")
);
LPOPERX WINAPI
xll_rand_err(void)
{
#pragma XLLEXPORT
	static OPERX o;

	o = ErrX(static_cast<WORD>(err[rand_between(0,sizeof(err))]));

	return &o;
}

static int type[] = {
	xltypeNum,
	xltypeStr,
	xltypeBool,
	xltypeErr,
	xltypeInt,
	xltypeMulti
};

LPOPERX WINAPI xll_rand_type(xword type);

static AddInX xai_rand_multi(
	_T("?xll_rand_multi"), XLL_LPOPERX XLL_VOLATILEX,
	_T("RAND.MULTI"), _T(""),
	_T("Test"), _T("Return a random two dimensional range. ")
);
LPOPERX WINAPI
xll_rand_multi(void)
{
#pragma XLLEXPORT
	static OPERX o;
	xword r, c;

	r = static_cast<xword>(1/rand_uniform());
	r = min(r, limits<XLOPERX>::maxrows);
	c = static_cast<xword>(1/rand_uniform());
	c = min(c, limits<XLOPERX>::maxcols);
	o.resize(r, c);
	xword t = static_cast<xword>(rand_between(0, dimof(type) - 1)); // not multi
	for (xword i = 0; i < o.size(); ++i)
		o[i] = *xll_rand_type(t);

	return &o;
}

static AddInX xai_rand_int(
	_T("?xll_rand_int"), XLL_DOUBLEX XLL_LONGX XLL_VOLATILEX,
	_T("RAND.INT"), _T("n"),
	_T("Test"), _T("Return a random integer in the range [0, n). ")
);
double WINAPI
xll_rand_int(LONG n)
{
#pragma XLLEXPORT

	return rand_between(0, n);
}

static AddInX xai_rand_type(
	_T("?xll_rand_type"), XLL_LPOPERX XLL_SHORTX XLL_VOLATILEX,
	_T("RAND.TYPE"), _T("Type"),
	_T("Test"), _T("Return a random OPER of the given Type. ")
);
LPOPERX WINAPI
xll_rand_type(xword type)
{
#pragma XLLEXPORT
	static OPERX o;

	if (type == xltypeNum)
		o = xll_rand_num();
	else if (type == xltypeStr) {
		const xchar* s = xll_rand_str();
		o = StrX(s + 1, s[0]);
	}
	else if (type == xltypeBool)
		o = (TRUE == xll_rand_bool());
	else if (type == xltypeInt)
		o = xll_rand_int(RAND_MAX);
	else if (type == xltypeMulti)
		o = *xll_rand_multi();
	else
		o = _T("");

	return &o;
}

#ifdef _DEBUG

bool
check_strict_weak(const OPERX& x, const OPERX& y, const OPERX& z)
{
	ensure (!(x < x));
	ensure (!(y < y));
	ensure (!(z < z));

	if (x != y && x < y)
		ensure (!(y < x));
	if (y != z && y < z)
		ensure (!(z < y));

	if (x < y && y < z)
		ensure (x < z);

	if (x < y)
		ensure (x < z || z < y);

	return true;
}

int
test_strict_weak()
{
	try {
		OPERX m = *xll_rand_multi();
		for (size_t i = 0; i + 2 < m.size(); ++i)
			check_strict_weak(m[i], m[i + 1], m[i + 2]);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

int
test_rand(void)
{
	test_strict_weak();

	return 1;
}
static Auto<OpenAfter> xao_test_rand(test_rand);

#endif // _DEBUG
