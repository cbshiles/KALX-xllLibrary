// test.cpp - Exercise the xll add-in library.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
// Uncomment to build for Excel 2007 and later.
#include "test.h"
#include "xll/hash.h"
#include "xll/xll.h"
//#include "document/document.h"

using namespace xll;

typedef traits<XLOPERX>::xstring xstring;

// Can't call MessageBox from xlAutoOpen. Well, you can...
inline void
Error(const char* err)
{
	XLOPER xErr;
	LPXLOPER pxErr[1];
	pxErr[0] = &xErr;
	xErr.xltype = xltypeStr;
	char n = static_cast<char>(strlen(err));
	xErr.val.str = new char[n + 1];
	xErr.val.str[0] = n;
	strncpy(xErr.val.str + 1, err, n);

	traits<XLOPER>::Excelv(xlcAlert, 0, 1, pxErr);

	delete [] xErr.val.str;
}

// Test in order of include files.

static bool xllb(true);
static int xlli(3);
static char xllc('c');
static double xlld(1.23);
static LPCTSTR xlls(_T("Foo"));

XLL_ENUM(xllb, XLLB, _T("Enum"), _T("Boolean true.")) 
XLL_ENUM(xlli, XLLI, _T("Enum"), _T("Integer 3.")) 
XLL_ENUM(xllc, XLLC, _T("Enum"), _T("Character 'c'.")) 
XLL_ENUM(xlld, XLLD, _T("Enum"), _T("Double 1.23.")) 
XLL_ENUM(xlls, XLLS, _T("Enum"), _T("String \"Foo\".")) 

void
test_ensure(void)
{
	bool caught(false);

	try {
		// fire ensure
		int i = 1;
		ensure (i != 1);
	}
	catch (const std::exception&) {
		caught = true;
	}

	ensure (caught);
}

void
test_defines(void)
{
	ensure (ExcelX(xlfEvaluate, StrX(_T("XLLB() = TRUE"))));
	ensure (ExcelX(xlfEvaluate, StrX(_T("XLLI() = 3"))));
	ensure (ExcelX(xlfEvaluate, StrX(_T("XLLC() = CODE(\"c\")"))));
	ensure (ExcelX(xlfEvaluate, StrX(_T("XLLD() = 1.23"))));
	ensure (ExcelX(xlfEvaluate, StrX(_T("XLLS() = \"foO\""))));

	LOPERX c = ExcelX(xlfEvaluate, StrX(_T("XLLC()")));
	ensure (c.xltype == xltypeNum);
	ensure (c.val.num == 'c');
}

#define N 30 // can work with more - empirically verified
void
test_traits(void)
{
	XLOPER x[N];

	for (int i = 0; i < N; ++i) {
		x[i].xltype = xltypeNum;
		x[i].val.num = i + 1;
	}

	XLOPER xResult;
	int result = traits<XLOPER>::Excel(xlfSum, &xResult, N
		,&x[0]
		,&x[1]
		,&x[2]
		,&x[3]
		,&x[4]
		,&x[5]
		,&x[6]
		,&x[7]
		,&x[8]
		,&x[9]
		,&x[10]
		,&x[11]
		,&x[12]
		,&x[13]
		,&x[14]
		,&x[15]
		,&x[16]
		,&x[17]
		,&x[18]
		,&x[19]
		,&x[20]
		,&x[21]
		,&x[22]
		,&x[23]
		,&x[24]
		,&x[25]
		,&x[26]
		,&x[27]
		,&x[28]
		,&x[29]
/*		,&x[30]
		,&x[31]
		,&x[32] // need to change traits<XLOPER>::argmax to test
		,&x[33]
		,&x[34]
		,&x[35]
		,&x[36]
		,&x[37]
		,&x[38]
		,&x[39]
*/
	);
	ensure (result == xlretSuccess);
	ensure (xResult.xltype == xltypeNum && xResult.val.num == N*(N+1)/2);
}

void
test_fp(void)
{
	FPX x;
	ensure (x.is_empty());

	FPX x2(1,2);
	ensure (!x2.is_empty());
	ensure (x2.rows() == 1);
	ensure (x2.columns() == 2);
	ensure (x2.size() == 2);

	x2[0] = 1;
	x2[1] = 2;
	ensure (x2(0, 0) == 1);
	ensure (x2(0, 1) == 2);

	x2.reshape(2,3);
	ensure (x2.rows() == 2);
	ensure (x2.columns() == 3);
	ensure (x2.size() == 6);
	ensure (x2(0, 0) == 1);
	ensure (x2(0, 1) == 2);

	FPX x3(x2);
	ensure (x3.rows() == 2);
	ensure (x3.columns() == 3);
	ensure (x3.size() == 6);
	ensure (x3(0, 0) == 1);
	ensure (x3(0, 1) == 2);

	FPX x4;
	x4 = x3;
	ensure (x4.rows() == 2);
	ensure (x4.columns() == 3);
	ensure (x4.size() == 6);
	ensure (x4(0, 0) == 1);
	ensure (x4(0, 1) == 2);
}

void
test_xloper(void)
{
	XLOPERX n;
	n.xltype = xltypeNum;
	n.val.num = 1.23;
	ensure (to_number(n) == 1.23);
	traits<XLOPERX>::xstring sn = to_string(n);
	ensure (sn.length() == 4);
	ensure (sn == _T("1.23"));

	XLOPERX s;
	s.xltype = xltypeStr;
	s.val.str = _T("\03foo");
	ensure (to_number(s) == 3); // !!!length
	ensure (to_string(s) == _T("foo"));

	// bool
	XLOPERX b;
	b.xltype = xltypeBool;
	b.val.xbool = true;
	ensure (to_number(b) == 1);
	xstring bs = to_string(b);
	size_t nn;
	nn = bs.length();
	ensure (to_string(b) == _T("TRUE"));

	// err
	XLOPERX e;
	e.xltype = xltypeErr;
	e.val.err = xlerrNA;

	XLOPERX m;
	m.xltype = xltypeMulti;
	m.val.array.rows = 2;
	m.val.array.columns = 3;
	XLOPERX lp[6];
	m.val.array.lparray = lp;
	for (size_t i = 0; i < 6; ++i) {
		m.val.array.lparray[i].xltype = xltypeNum;
		m.val.array.lparray[i].val.num = i;
	}
	ensure (to_string(m) == _T("012345"));
	ensure (to_string(m, _T(", ")) == _T("0, 1, 23, 4, 5"));
	ensure (to_string(m, _T(", "), _T(";")) == _T("0, 1, 2;3, 4, 5"));

	ensure (rows(n) == 1);
	ensure (rows(s) == 1);
	ensure (rows(b) == 1);
	ensure (rows(e) == 1);
	ensure (rows(m) == 2);

	ensure (columns(n) == 1);
	ensure (columns(b) == 1);
	ensure (columns(e) == 1);
	ensure (columns(s) == 1);
	ensure (columns(m) == 3);

	ensure (size(n) == 1);
	ensure (size(b) == 1);
	ensure (size(e) == 1);
	ensure (size(s) == 1);
	ensure (size(m) == 6);

	ensure (index(m, 4).val.num == 4);
	ensure (index(m, 1, 1).val.num == 4);
	index(m, 1, 1) = s; // default XLOPER::operator=()
	ensure (index(m, 1, 1).xltype == s.xltype);
	ensure (index(m, 1, 1).val.str[0] == s.val.str[0]);

	ensure (begin(n) == &n);
	ensure (begin(s) == &s);
	ensure (begin(m) == m.val.array.lparray);

	ensure (end(n) == &n + 1);
	ensure (end(s) == &s + 1);
	ensure (end(m) == m.val.array.lparray + size(m));

	ensure (n == n);
	ensure (!(n != n));
	ensure (!(n < n));
	ensure (n <= n);
	ensure (!(n > n));
	ensure (n >= n);
	ensure (n < s);
	ensure (s == s);
	ensure (s < b);
	ensure (b == b);
	ensure (b < e);
	ensure (e == e);
	ensure (e < m);
	ensure (m == m);
}

void
test_oper(void)
{
	OPERX o; // OPER()
	ensure (o.xltype == xltypeNil);
	ensure (!o);

	OPERX o1(1.23); // OPER(double)
	ensure (o1.xltype == xltypeNum);
	ensure (o1.val.num == 1.23);
	ensure (o1);
	ensure (o1 == 1.23); // operator double() const
	ensure (1.23 == o1); // same
	o1 = 3.45; // operator=(double)
	ensure (o1 == 3.45);
	ensure (3.45 == o1);

	// watch this, ma!
	o1 = o1 + 1.23;
	ensure (o1 == 3.45 + 1.23);
	o1 = 1.23 - o1;
	ensure (o1 == 1.23 - (3.45 + 1.23));
	// o1 -= 1.23; // this won't work

	ensure (!OPERX(_T("")));

	OPERX o2(_T("foo")); // OPER(LPCTSTR)
	ensure (o2.xltype == xltypeStr);
	ensure (0 == _tcsnicmp(o2.val.str + 1, _T("fOo"), o2.val.str[0]));
	ensure (o2);

	OPERX o3(o2.val.str + 1, o2.val.str[0]); // OPER(LPCTSTR, TCHAR)
	ensure (o3.xltype == xltypeStr);
	ensure (0 == _tcsnicmp(o3.val.str + 1, _T("fOo"), o3.val.str[0]));

	o3 = _T("barbaz");
	ensure (o3.xltype == xltypeStr);
	ensure (o3 == _T("BaRBaZ"));

	OPERX o4(true);
	ensure (o4.xltype == xltypeBool);
	ensure (o4.val.xbool == TRUE);
	ensure (o4);
	ensure (o4 == true);

	BoolX b;
	ensure (b.xltype == xltypeBool);
	ensure (b.val.xbool == FALSE);
	ensure (b.operator==(false));
	ensure (!b);

	IntX i;
	ensure (i.xltype == xltypeInt);
	ensure (i.val.w == 0);
	ensure (i == 0);

	i = 4; // you might think it is an int
	ensure (i.xltype == xltypeNum); // but it is not, just like Excel
	ensure (i.val.num == 4);
	ensure (i == 4);

	ErrX err(xlerrNA);
	ensure (err.xltype == xltypeErr);
	ensure (err.val.err == xlerrNA);

	ensure (MissingX().xltype == xltypeMissing);
	ensure (NilX().xltype == xltypeNil);
}

void
test_num(void)
{
	OPERX n(1);
	ensure (n.xltype == xltypeNum);
	ensure (n.val.num == 1);

	n = n + n*n - 1/n + 2; 
	ensure (3 == n);

	OPERX n2(NumX(1.2));
	ensure (n2.xltype == xltypeNum);
	ensure (n2.val.num == 1.2);

	NumX num;
	ensure (num.xltype == xltypeNum);
	ensure (num.val.num == 0);
	ensure (num == 0); 
	ensure (0. == num); // 0 == num is ambiguous

	NumX num2(1.23);
	ensure (num2.xltype == xltypeNum);
	ensure (num2.val.num == 1.23);
	ensure (num2 == 1.23);
	ensure (1.23 == num2);
	num2 = 3.45;
	ensure (num2 == 3.45);

	// NumX num(o1); // will not compile

	num2 = OPERX(); // o1 is a num
	ensure (num2.xltype == xltypeNum);
	ensure (num2.val.num == 0);
	ensure (num2 == 0);

	NumX num3;
	num3 = StrX(_T("foo")); // operator double() on o3 which is xltypeStr
	ensure (num3.xltype == xltypeNum);
	ensure (num3 == 3); // string length!

	NumX num4(BoolX(true)); // operator double() on o4 which is xltpeBool
	ensure (num4.xltype == xltypeNum);
	ensure (num4.val.num == 1);
	ensure (num4);
	ensure (num4 == 1);
}

void
test_str(void)
{
	StrX str;
	ensure (str.xltype == xltypeStr);
	ensure (str.val.str[0] == 0);
	ensure (!str); // operator double()

	StrX str2(_T("hello"));
	ensure (str2.xltype == xltypeStr);
	ensure (0 == _tcsncmp(str2.val.str + 1, _T("hello"), str2.val.str[0]));
	ensure (str2 == _T("HeLlO"));

	StrX str3(str2);
	ensure (str3.xltype == xltypeStr);
	ensure (0 == _tcsncmp(str3.val.str + 1, _T("hello"), str3.val.str[0]));

	// str = o1; // won't compile
	// StrX str(o1); // won't compile

	str3 = _T("bye");
	ensure (str3 == _T("Bye"));

	// str3 = NumX(1.23); // won't compile???
	// str3(NumX(1.23); // won't compile???
}

void
test_multi(void)
{
	OPERX m(1, 3);
	for (size_t i = 0; i < m.size(); ++i)
		m[i] = i;
	m.push_back(NumX(3));
	m.push_back(NumX(4));
	m.push_back(NumX(5));
	ensure (m[5] == 5);
	m.resize(2, 3);

	OPERX m2(1, 3);
	for (size_t i = 0; i < m2.size(); ++i)
		m2[i] = m.size() + i;

	m.push_back(m2);
	ensure (m.rows() == 3);
	ensure (m.columns() == 3);
	for (size_t i = 0; i < m.size(); ++i)
		ensure (m[i] == i);

	m(1,2) = OPERX(_T("blah"));
	ensure (m(1, 2) == _T("BlAh"));
	m(1,2) = StrX(_T("derp"));
	ensure (m(1, 2) == _T("DerP"));

	for (LPOPERX p = m.begin(); p != m.end(); ++p)
		ensure (p->xltype);

	OPERX x;
	// calls operator=()
	x.push_back(StrX(_T("String")));
	ensure (x.xltype == xltypeStr);
	ensure (x == _T("sTRING"));

	// preserves existing values
	x.push_back(NumX(1.23));
	ensure (x.xltype == xltypeMulti);
	ensure (x.rows() == 2);
	ensure (x.columns() == 1);
	ensure (x[0] == _T("sTrInG"));
	ensure (x[1] == 1.23);
}

void
test_oper_nice(void)
{
	// abstract data type tests
	OPERX o;
	OPERX n(1.23);
	OPERX s(_T("bar"));
	OPERX m(2, 3);

	o = n;
	ensure (o.xltype == xltypeNum);
	ensure (o == 1.23);

	o = s;
	ensure (o.xltype == xltypeStr);
	ensure (o == _T("bAr"));

	o = m;
	ensure (o[3].xltype == xltypeNil);
	o[3] = n;
	ensure (o[3] == n);
	o[4] = s;
	ensure (o(1,1) == s);

	OPERX m2(m);
	ensure (m2.xltype == xltypeMulti);
	ensure (m2.rows() == m.rows());
	ensure (m2.columns() == m.columns());
	ensure (m2[4] == m[4]);
}

void
test_oper_relops(void)
{
	OPERX o;
	ensure (o == o);
	ensure (!(o != o));
	ensure (!(o < o));
	ensure (o <= o);
	ensure (!(o > o));
	ensure (o >= o);

	ensure (OPERX(1.23) == OPERX(1.23));
	ensure (OPERX(1.23) < OPERX(2.34));

	ensure (OPERX(1.23) < OPERX(_T("baz")));
	ensure (OPERX(_T("bar")) == OPERX(_T("bAr")));
	ensure (OPERX(_T("bar")) < OPERX(_T("baZ")));
	ensure (OPERX(_T("bar")) < OPERX(_T("baRz")));

	ensure (OPERX(_T("true")) < OPERX(false));
	ensure (OPERX(false) == OPERX(false));
	ensure (OPERX(false) < OPERX(true));
}

void
test_addin(void)
{
	return;
}

void
test_hash(void)
{
	unsigned int ui, ui2;

	ui = utility::hash("hash");
	ui2 = utility::hash(L"hash");

	ensure (ui == ui2);
}

void
test_find(void)
{
	XLOPERX xId = ExcelX(xlfEvaluate, StrX(_T("XLL.FOO")));

	ArgsMapX::args_map& m(ArgsMapX::Map());
	const ArgsX* pfoo;
	pfoo = m[xId.val.num];

	ensure (pfoo->FunctionText() == _T("XLL.FOO"));
}

void
test_handle(void)
{
	handle<std::string> h(new std::string());

	HANDLEX hx = h.handlex();

	handle<std::string> k(hx);
	k->operator=("hi");

	ensure (k->compare("hi") == 0);
}

int
xll_test(void)
{
	try {
		test_defines();
		test_ensure();
		test_traits();
		test_fp();
		test_xloper();
		test_oper();
		test_num();
		test_str();
		test_multi();
		test_oper_nice();
		test_oper_relops();

		test_hash();
		test_find();

		test_handle();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

static Auto<OpenAfter> xao_test(xll_test);

// register macro function
int WINAPI 
xll_macro(void)
{
#pragma XLLEXPORT

	StrX xAlert(_T("Alert"));

	traits<XLOPERX>::Excel(xlcAlert, 0, 1, &xAlert);

	return 1;
}
int
xai_macro(void)
{
	AddInRegister<XLOPERX>(_T("?xll_macro"), _T("XLL.MACRO"));

	return 1;
}
static Auto<OpenAfter> xao_macro(xai_macro);

// register worksheet function
double WINAPI
xll_function(double num)
{
#pragma XLLEXPORT

	return 2*num;
}
int
xai_function(void)
{
	AddInRegister<XLOPERX>(_T("?xll_function"), XLL_DOUBLEX XLL_DOUBLEX, _T("XLL.FUNCTION"), _T("Num"));

	return 1;
}
static Auto<OpenAfter> xao_function(xai_function);

// register worksheet function mark 2
double WINAPI
xll_function2(double num)
{
#pragma XLLEXPORT

	return 2*num;
}
int
xai_function2(void)
{
	AddInRegister<XLOPERX>(
		ArgsX(_T("?xll_function2"), XLL_DOUBLEX XLL_DOUBLEX, _T("XLL.FUNCTION2"), _T("Num"))
		.Category(_T("My Category"))
		.FunctionHelp(_T("This function doubles it's argument."))
		.ArgumentHelp(_T("is a number. "))
	);

	return 1;
}
static Auto<OpenAfter> xao_function2(xai_function2);

// test LOPER class
LPXLOPERX WINAPI
xll_loper(void)
{
#pragma XLLEXPORT
	static LOPERX o;

	o = ExcelX(xlGetName);

	return &o;
}
int
xai_loper(void)
{
	AddInRegister<XLOPERX>(_T("?xll_loper"), XLL_LPXLOPERX, _T("XLL.LOPER"), _T(""));

	return 1;
}
static Auto<OpenAfter> xao_loper(xai_loper);

// The right way to do it.
static AddInX xai_sqrt(
	_T("?xll_sqrt"), XLL_LPXLOPERX XLL_DOUBLEX,
	_T("XLL.SQRT"), _T("Num")
);
LPXLOPERX WINAPI
xll_sqrt(double num)
{
#pragma XLLEXPORT
	static OPERX x;

	if (num < 0)
		x = ErrX(xlerrNum); // #NUM!
	else
		x = sqrt(num);

	return &x;
}

static AddInX xai_macro2(_T("?xll_macro2"), _T("XLL.MACRO2"));
int WINAPI
xll_macro2(void)
{
#pragma XLLEXPORT

	return 1;
}

// The new way to do it.
static AddInX xai_foo(ArgsX(XLL_DOUBLEX, _T("?xll_foo"), _T("XLL.FOO")));
double WINAPI
xll_foo(void)
{
#pragma XLLEXPORT

	return 37337;
}

static AddInX xai_foo2(
	ArgsX(XLL_DOUBLEX, _T("?xll_foo2"), _T("XLL.FOO2"))
	.Arg(XLL_CSTRINGX, _T("Str"), _T("is a string "))
);
double WINAPI
xll_foo2(LPCTSTR str)
{
#pragma XLLEXPORT

	return _tcsclen(str);
}

static AddInX xai_foo3(
	ArgsX(XLL_LPOPERX, _T("?xll_foo3"), _T("XLL.FOO3"))
	.Arg(XLL_CSTRINGX, _T("Str"), _T("is a string"))
	.Arg(XLL_BOOLX, _T("Bool"), _T("is a boolean "))
	.Category(_T("Foo Category"))
	.FunctionHelp(_T("Returns string or boolean."))
);
LPOPERX WINAPI
xll_foo3(LPCTSTR str, BOOL b)
{
#pragma XLLEXPORT
	static OPERX o;

	if (b)
		o = true;
	else
		o = str;

	return &o;
}

// Documentation
static AddInX xai_foo4(
	FunctionX(XLL_DOUBLEX, _T("?xll_foo4"), _T("XLL.FOO4"))
	.Arg(XLL_DOUBLEX, _T("Num"), _T("is a number. "))
	.Category(_T("Foo Category"))
	.FunctionHelp(_T("A function that does foo."))
/*	.Documentation(xml::element()
		.content(xml::element(_T("para")).content(_T("This is a paragraph.")))
		.content(xml::element(_T("para")).content(_T("And this is another.")))
	)
*/
);
double WINAPI
xll_foo4(double num)
{
#pragma XLLEXPORT
	XLL_ERROR("hi");

	return 2*num;
}
