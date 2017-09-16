// test.cpp - Exercise the xll add-in library.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
// Uncomment to build for Excel 2007 and later.
#include "test.h"
/*
#include "xll/utility/hash.h"
#include "xll/xll.h"
#include "xll/maml/document.h"
*/
using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xfp xfp;
typedef traits<XLOPERX>::xword xword;
typedef traits<XLOPERX>::xstring xstring;

// Can't call MessageBox from xlAutoOpen. Well, you can...
inline void
xll_error(const char* err)
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
	ensure (ExcelX(xlfEvaluate, OPERX(_T("XLLB()"))));
	ensure (ExcelX(xlfEvaluate, OPERX(_T("XLLI() = 3"))));
	ensure (ExcelX(xlfEvaluate, OPERX(_T("XLLC() = CODE(\"c\")"))));
	ensure (ExcelX(xlfEvaluate, OPERX(_T("XLLD() = 1.23"))));
	ensure (ExcelX(xlfEvaluate, OPERX(_T("XLLS() = \"foO\""))));

	LOPERX c = ExcelX(xlfEvaluate, OPERX(_T("XLLC()")));
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
	// cyclic indices
	ensure (x2[2] == x2[0]);
	ensure (x2[3] == x2[1]);

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

	FPX x5(2, 3);
	for (xword i = 0; i < x5.size(); ++i)
		x5[i] = i;
	ensure (x5(1, 2) == 5);

	FPX x6(2, 3);
	x6.set(3, 2, x5.array());
	ensure (x6(2, 1) == 5);

	FPX x7(1, 3);
	x7.push_back(x7.begin(), x7.end());
	ensure (x7.rows() == 1);
	ensure (x7.columns() == 6);
	x7.push_back(x7.begin(), x7.end(), true);
	ensure (x7.rows() == 2);
	ensure (x7.columns() == 6);

	OPERX o7(1, 6);
	o7[0] = 1.23;
	x7.push_back(o7.begin(), o7.end()); // this can works!
	ensure (x7.rows() == 3);
	ensure (x7.columns() == 6);
	ensure (x7(2, 0) == 1.23);

	FPX x8(3, 1);
	x8.push_back(x8.begin(), x8.end());
	ensure (x8.rows() == 6);
	ensure (x8.columns() == 1);

	x8.push_back(x7.begin(), x7.begin() + 2);
	ensure (x8.rows() == 8);
	ensure (x8.columns() == 1);

	xfp* px8 = x8.get();
	ensure (px8->rows == x8.rows());
	ensure (px8->columns == x8.columns());
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
	ensure (bs == _T("TRUE"));

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
	for (xword i = 0; i < 6; ++i) {
		m.val.array.lparray[i].xltype = xltypeNum;
		m.val.array.lparray[i].val.num = i;
	}
	ensure (to_string(m) == _T("012345"));
	ensure (to_string(m, _T(", ")) == _T("0, 1, 23, 4, 5"));
	ensure (to_string(m, _T(", "), _T(";")) == _T("0, 1, 2;3, 4, 5"));

	ensure (to_string<XLOPERX>(ErrX(xlerrNA)) == _T("#N/A"));

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
test_loper(void)
{
	LOPERX o;
	ensure (!o.owner);

	LOPERX o2(o);
	ensure (!o2.owner);
	
	LOPERX o3(o, true);
	ensure (o3.owner);
	o3.owner = false; // don't want xlFree called

	LOPERX* po4;
	{
		LOPERX o4 = ExcelX(xlfText, ExcelX(xlfNow), OPERX(_T("yyyy")));
		// str != num
		ensure (o4 != ExcelX(xlfYear, ExcelX(xlfNow)));
		ensure (ExcelX(xlfValue, o4) == ExcelX(xlfYear, ExcelX(xlfNow)));
		ensure (o4.owner);
		po4 = (LOPERX*)&o4; // operator& returns LPXLOPER
	}
	ensure (!po4->owner);

	OPERX o5 = ExcelX(xlfLen, ExcelX(xlfNow));
	ensure (o5.xltype == xltypeNum);

}

void
test_oper(void)
{
	ensure (OPERX().size() == 0);
	ensure (size<XLOPERX>(MissingX()) == 0);
	ensure (size<XLOPERX>(NilX()) == 0);

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

	OPERX m0(1, 1);
	m0[0] = 1.23;
	ensure (m0 == 1.23);
	ensure (m0 != 3.21);

	OPERX m2(1, 2);
	ensure (m0 != 0);

	ErrX err(xlerrNA);
	ensure (err.xltype == xltypeErr);
	ensure (err.val.err == xlerrNA);

	ensure (MissingX().xltype == xltypeMissing);
	ensure (NilX().xltype == xltypeNil);

	OPERX m3(1, 2), m4(1, 2);
	m3[0] = _T("2");
	m4[0] = _T("1");

	ensure (m3[1].xltype == xltypeNil);
	ensure (m4[1].xltype == xltypeNil);
	ensure (m4 < m3);
	
	m3[0] = _T("A");
	m3[1] = 2;
	m4[0] = _T("A");
	m4[1] = 1;
	ensure (m3 > m4);

	m4.resize(0,0);
	ensure (m4.size() == 0);

	m4.push_back(OPERX(_T("foo")));
	ensure (m4.size() == 1);
	ensure (m4.xltype == xltypeStr);
	ensure (m4 == _T("foo"));

	m4.push_back(OPERX(1.23));
	ensure (m4.size() == 2);
	ensure (m4.xltype == xltypeMulti);
	ensure (m4[1] == 1.23);

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

	str3.append(_T("bye"));
	ensure (str3 == _T("byebye"));

	str3.append(_T("birdie"), 4);
	ensure (str3 == _T("byebyebird"));

	str3 = _T("");
	ensure (str3.val.str[0] == 0);
}


OPERX*
my_helper(const OPERX& o)
{
	static OPERX x;

	x = o;
	x = _T("foo");

	return &x;
}
void
test_multi(void)
{
	OPERX m(1, 3);
	for (xword i = 0; i < m.size(); ++i)
		m[i] = i;
	m.push_back(NumX(3));
	m.push_back(NumX(4));
	m.push_back(NumX(5));
	ensure (m[5] == 5);
	m.resize(2, 3);

	OPERX m2(1, 3);
	for (xword i = 0; i < m2.size(); ++i)
		m2[i] = m.size() + i;

	m.push_back(m2);
	ensure (m.rows() == 3);
	ensure (m.columns() == 3);
	for (xword i = 0; i < m.size(); ++i)
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

	ensure(x);
	x[0] = true;
	x[1] = false;
	ensure (!x); // not all entries are true

	x.resize(1, 1);
	ensure (x);
	x[0] = false;
	ensure (!x);

	{
		// trigger o.~OPER()
		OPERX o;
		o.resize(1, 2);
		o[0] = _T("foo");
		o[1] = _T("bar");
		o.resize(1, 3);
		o[0] = _T("baz");
		o[0] = StrX(_T("blah"));
		o[2] = o[0];

		o[1] = *my_helper(StrX(_T("foo")));
		o.resize(4, 1);
		o[1] = *my_helper(StrX(_T("bar")));
		o[2] = *my_helper(OPERX(_T("baz")));

		o.resize(3,5);
//		ensure (o[1] == _T("bar"));
//		ensure (o[2] == _T("baz"));
		o[10] = _T("blah");
	}

	x.resize(2,3);
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
test_bigdata(void)
{
	xchar* s = new xchar[2];

	s[0] = _T('a');
	s[1] = 0;
	OPERX b(2, s);
	delete[] s;
	ensure (b.xltype == xltypeBigData);
	ensure (b.val.bigdata.cbData == 2);
	ensure (b.val.bigdata.h.lpbData[0] == _T('a'));
	ensure (b.val.bigdata.h.lpbData[1] == 0);

	OPERX b2(b);
	ensure (b2.xltype == xltypeBigData);
	ensure (b2.val.bigdata.cbData == 2);
	ensure (b2.val.bigdata.h.lpbData[0] == _T('a'));
	ensure (b2.val.bigdata.h.lpbData[1] == 0);

	OPERX b3;
	b3 = b2;
	ensure (b3.xltype == xltypeBigData);
	ensure (b3.val.bigdata.cbData == 2);
	ensure (b3.val.bigdata.h.lpbData[0] == _T('a'));
	ensure (b3.val.bigdata.h.lpbData[1] == 0);

	b3.val.bigdata.h.lpbData[0] = _T('b');
	b2 = b3;
	ensure (b2.xltype == xltypeBigData);
	ensure (b2.val.bigdata.cbData == 2);
	ensure (b2.val.bigdata.h.lpbData[0] == _T('b'));
	ensure (b2.val.bigdata.h.lpbData[1] == 0);

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
	LOPERX xId = ExcelX(xlfEvaluate, StrX(_T("XLL.FOO")));

	ArgsMapX::args_map& m(ArgsMapX::Map());
	const ArgsX* pfoo;
	pfoo = m[xId.val.num];

	ensure (pfoo->FunctionText() == _T("XLL.FOO"));
}

void
test_handle(void)
{
	handle<std::string> h(new std::string());

	HANDLEX hx = h.get();

	handle<std::string> k(hx);
	k->operator=("hi");

	ensure (k->compare("hi") == 0);
}

void
test_excelv(void)
{
	OPERX o(3, 1);

	o[0] = 1;
	o[1] = 2;
	o[2] = 3;

	OPERX o1 = Excelv<XLOPERX>(xlfSum, o.size(), &o[0]);
	ensure (o1.xltype == xltypeNum);
	ensure (o1 == 6);
}

void
test_mref(void)
{
	MREFX r;
	OPERX o(1, 2, 3, 4);
	r[0] = o.val.sref.ref;
	ensure (r.val.mref.lpmref->count == 1);
	ensure (r.val.mref.lpmref->reftbl[0].rwFirst == 1);
	ensure (r.val.mref.lpmref->reftbl[0].colFirst == 2);
	ensure (r.val.mref.lpmref->reftbl[0].rwLast == 3);
	ensure (r.val.mref.lpmref->reftbl[0].colLast == 5);
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
		test_loper();
		test_oper();
		test_num();
		test_str();
		test_multi();
		test_oper_nice();
		test_oper_relops();
		test_bigdata();
		test_mref();

		test_hash();
		test_find();

		test_excelv();

//		test_handle();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

static Auto<OpenAfterX> xao_test(xll_test);

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
	Register<XLOPERX>(_T("?xll_macro"), _T("XLL.MACRO"));

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
	Register<XLOPERX>(_T("?xll_function"), XLL_DOUBLEX XLL_DOUBLEX, _T("XLL.FUNCTION"), _T("Num"));

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
	Register<XLOPERX>(
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
	Register<XLOPERX>(_T("?xll_loper"), XLL_LPXLOPERX, _T("XLL.LOPER"), _T(""));

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
	.Alias(_T("XLL.FOOTWO"))
);
double WINAPI
xll_foo2(LPCTSTR str)
{
#pragma XLLEXPORT

	return _tcsclen(str);
}

using xml::element;

#define CATEGORY _T("Foo Category")
/*
static AddInX xai_foo_doc(
	ArgsX(CATEGORY)
	.Documentation(xml::File(_T("cuda.xml")))
);
*/
static AddInX xai_foo_doc(
	ArgsX(CATEGORY)
	.Documentation(_T("This is some documentation."))
);

static AddInX xai_foo3(
	ArgsX(XLL_LPOPERX, _T("?xll_foo3"), _T("XLL.FOO3")) 
	.Arg(XLL_CSTRINGX, _T("Str"), _T("is a string"))
	.Arg(XLL_BOOLX, _T("Bool"), _T("is a boolean "))
	.Sort(_T("2")) // after foo4
	.Category(_T("Foo Category"))
	.FunctionHelp(_T("Returns string or boolean."))
	.Documentation(_T("Foo 3 documentation"))
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

/*
MATH(VAR(e) SUP(ENT_pi VAR(i)) _T(" = - 1"))
*/
using namespace xml;

#define V_(v) _T("<markup><i>") e _T("</i></markup>")
// Documentation
static AddInX xai_foo4(
	FunctionX(XLL_DOUBLEX, _T("?xll_foo4"), _T("XLL.FOO4") _T("\0") _T("1")) 
	.Arg(XLL_DOUBLEX, _T("Num"), _T("is a number. "))
	.Sort(_T("1"))  // before foo3
	.Category(CATEGORY)
	.FunctionHelp(_T("A function that does foo."))
	.Documentation(
	// e^{pi i} = -1
		math(element()
			._( _T("<markup><i>") _T("e") _T("</i></markup>") )
			._(sup(_T("<markup><i>") ENT_pi _T("i") _T("</i></markup>")))
			._(_T(" = -1"))
		)
		, // see also
		element()
		.content(xlink(_T("XLL.FOO3")))
		.content(externalLink(_T("google"), _T("http://google.com/")))
	)
);
double WINAPI
xll_foo4(double num)
{
#pragma XLLEXPORT
	XLL_ERROR("hi");

	bool b;
	b = xai_foo4.Arg().Sort() < xai_foo3.Arg().Sort();

	return 2*num;
}
// macro documentation

static AddInX xai_macro_doc(
	MacroX(_T("?xll_macro_doc"), _T("MACRO.DOC"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Description of macro. "))
	.Documentation(_T("This is the macro documentation. "))
);
int WINAPI
xll_macro_doc(void)
{
#pragma XLLEXPORT

	return 1;
}


using namespace xml;

void
test_element(void)
{
	element t(_T("xml"));
	t.attribute(_T("version"), _T("1.0"));
	t.content(_T("An xml doc."));
	std::basic_string<TCHAR> s;
	s = t;
	ensure (s == _T("<xml version=\"1.0\">An xml doc.</xml>\n"));

	t.attribute(_T("level"), _T("3"));
	s = t;
	ensure (s == _T("<xml version=\"1.0\" level=\"3\">An xml doc.</xml>\n"));

	s = element(_T("xml"))
		.attribute(_T("version"), _T("2.0"))
		.content(_T("An xml doc."));
	ensure (s == _T("<xml version=\"2.0\">An xml doc.</xml>\n"));
	s = element(_T("tag"));
	ensure (s == _T("<tag />\n"));

	element esc(_T("esc"));
	esc.content(element::escape(_T("<>&\'\"")));
	s = esc;
	ensure (s == _T("<esc>&lt;&gt;&amp;&apos;&quot;</esc>\n"));

	s = element()
		.content(element(_T("tag")).attribute(_T("attr"), _T("value")).content(_T("text")))
		;
	ensure (s == _T("<tag attr=\"value\">text</tag>\n"));
	s =		element(_T("html"))
			.content(
				element(_T("head")).content(_T("H")))
			.content(
				element(_T("body")).content(_T("B")));
	ensure (s == _T("<html><head>H</head>\n<body>B</body>\n</html>\n"));
	s =		element()
			.content(
				element(_T("head")).content(_T("H")))
			.content(
				element(_T("body")).content(_T("B")));
	ensure (s == _T("<head>H</head>\n<body>B</body>\n"));

	element e(_T("tag"));
	e._(_T("attr"), _T("val"));
	e._(_T("attr2"), _T("val2"));
	e._(_T("inner text"));
	e._(_T("more text"));
	ensure (0 == _tcscmp(e, _T("<tag attr=\"val\" attr2=\"val2\">inner textmore text</tag>\n")));

	para p(_T("this is paragraph text"));
	ensure (0 == _tcscmp(p , _T("<para>this is paragraph text</para>\n")));

	typedef xml::element _;

	math m; // e^{pi i} = -1
	m._(var(_T("e")))
	 ._( sup( _()._(ENT_pi)._(var(_T("i")))) )
	 ._(_T(" = -1"));
	s = m;
}

void
test_conceptual(void)
{
	std::basic_string<TCHAR> s;

	Conceptual c1(_T("intro"));
	s = c1;
	std::basic_string<TCHAR> s1(_T("<introduction><para>intro</para>\n</introduction>\n"));
	ensure (s == s1);

	Conceptual c2(_T("intro"), _T("sum"));
	s = c2;
	std::basic_string<TCHAR> s2(_T("<summary><para>sum</para>\n</summary>\n"));
	s2 += s1;
	ensure (s == s2);

	c2.section(_T("title"), _T("content"));
	s = c2;
	std::basic_string<TCHAR> s3(s2);
	s3 += _T("<section><title>title</title>\n<content><para>content</para>\n</content>\n</section>\n");
	ensure (s == s3);

	c2.relatedTopics(element().content(_T("text")));
	s = c2;
	std::basic_string<TCHAR> s4(s3);
	s4 += _T("<relatedTopics>text</relatedTopics>\n");
	ensure (s == s4);

/*
	element t(element().content(
		table(_T("Table Title"))
		.header()
			.entry(_T("H1"))
			.entry(_T("H2"))
		.row()
			.entry(_T("k1"))
			.entry(_T("k2"))
		.row()
			.entry(_T("l1"))
			.entry(_T("l2"))
	));
*/
}

int
test_doc(void)
{
	try {
		test_element();
		test_conceptual();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfter> xao_test_doc(test_doc);

static AddInX xai_hi_macro(_T("?xll_hi_macro"), _T("HI.MACRO"));
int WINAPI
xll_hi_macro(void)
{
#pragma XLLEXPORT

	ExcelX(xlcAlert, OPERX(_T("Hi!")));

	return 1;
}
// On<XXX> execute macro
//static On<Key> xok_hi_macro(ON_CTRL ON_SHIFT "#", "HI.MACRO");

int
xll_test_eval()
{
	try {
		OPERX o;

		o = OPERX(_T("1.23"));
		o = ExcelX(xlfEvaluate, o);
		ensure (o == 1.23);

		o = OPERX(_T("{1.23,4.56}"));
		o = ExcelX(xlfEvaluate, o);
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 1);
		ensure (o.columns() == 2);
		ensure (o(0, 0) == 1.23);
		ensure (o(0, 1) == 4.56);

		o = OPERX(_T("{1.23;\"abc\"}"));
		o = ExcelX(xlfEvaluate, o);
		ensure (o.xltype == xltypeMulti);
		ensure (o.rows() == 2);
		ensure (o.columns() == 1);
		ensure (o(0, 0) == 1.23);
		ensure (o(1, 0) == _T("abc"));

		o = OPERX(_T("#N/A"));
		o = ExcelX(xlfEvaluate, o);
		ensure (o.xltype == xltypeErr);
		ensure (o.val.err == xlerrNA);
		// find list of all error strings!!!
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfter> xao_eval(xll_test_eval);
