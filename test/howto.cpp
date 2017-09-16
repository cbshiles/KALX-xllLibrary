// howto.cpp - How To Build an Add-in (XLL) for Excel Using Visual C++
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
// Compare with http://support.microsoft.com/kb/178474
// Uncomment to build for Excel 2007 and later.
//#define EXCEL12
#include <sstream>
#include "xll/xll.h"

#define CATEGORY _T("MyCat")

using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xcstr xcstr;

int
xll_aomb(void)
{
	ExcelX(xlcAlert, OPERX(_T("xlAutoOpen() called!")));

	return 1;
}
static Auto<Open> xao_aomb(xll_aomb);

int
xll_acmb(void)
{
	ExcelX(xlcAlert, OPERX(_T("xlAutoClose()")));

	return 1;
}
static Auto<Close> xac_acmb(xll_acmb);

// menu
int
xll_add_menu(void)
{
	OPERX m(2, 5);

	m(0,0) = _T("&MyMenu"); m(0,3) = _T("Joe's Xll menu!!!");
	m(1,0) = _T("M.O.T.D."); m(1,1) = _T("MyMotd"); m(1, 3) = _T("Message of the Day!");

	return ExcelX(xlfAddMenu, OPERX(1), m, OPERX(_T("Help"))) ? 1 : 0;
}
static Auto<Open> xao_add_menu(xll_add_menu);
int
xll_delete_menu(void)
{
	return ExcelX(xlfDeleteMenu, OPERX(1), OPERX(_T("MyMenu"))) ? 1 : 0;
}
static Auto<Close> xac_delete_menu(xll_delete_menu);

static AddInX xai_mymotd(_T("?xll_mymotd"), _T("MyMotd"));
int WINAPI
xll_mymotd(void)
{
#pragma XLLEXPORT

	static xcstr name[] = {
		_T("Rebekah"),
		_T("Brent"),
		_T("John"),
		_T("Joseph"),
		_T("Robert"),
		_T("Sara")
	};
	static xcstr quote[] = {
		_T("An apple a day, keeps the doctor away!"),
		_T("Carpe Diem: Seize the Day!"),
		_T("What you dare to dream, dare to do!"),
		_T("I think, therefore I am."),
		_T("A place for everything, and everything in its place."),
		_T("Home is where the heart is.")
	};

	if (!ExcelX(xlcAlert, ExcelX(xlfConcatenate, OPERX(name[rand()%dimof(name)]), OPERX(_T(" says ")), OPERX(quote[rand()%dimof(quote)]))))
		return 0;

	return 1;
}

static AddInX xai_myfunc(
	_T("?xll_myfunc"), XLL_LONGX XLL_LONGX XLL_LONGX,
	_T("MyFunc"), _T("Param1, Param2"),
	CATEGORY, _T("Example function that returns the product of its two parameters.")
);
long WINAPI
xll_myfunc(long parm1, long parm2) 
{
#pragma XLLEXPORT

	std::basic_ostringstream<xchar> buf;

	buf <<  _T("You sent ") << parm1 << _T(" and ") << parm2 << _T(" to MyFunc()!");
	::MessageBox(NULL, buf.str().c_str(), _T("MyFunc() in Anewxll!!!"), MB_SETFOREGROUND); // Can't call xlcAlert from function.

	return parm1 * parm2;
}
